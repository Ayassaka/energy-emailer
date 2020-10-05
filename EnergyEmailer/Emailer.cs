using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Net.Mail;
using System.Net;

namespace EnergyEmailer
{
    public static class Emailer
    {
        //private const string MESSAGE_CONTROL = @"html\ReportCard_Control.html";
        //private const string MESSAGE_GENERIC = @"html\ReportCard_Generic.html";
        //private const string MESSAGE_PERSONALIZED = @"html\ReportCard_Personalized.html";


        //public const int MESSAGE_TYPE_CONTROL = 0;
        //public const int MESSAGE_TYPE_GENERIC = 1;
        //public const int MESSAGE_TYPE_PERSONALIZED = 2;

        //private static EmailBody s_emailBodyControl = new EmailBody(MESSAGE_CONTROL);
        //private static EmailBody s_emailBodyGeneric = new EmailBody(MESSAGE_GENERIC);
        //private static EmailBody s_emailBodyPersonalized = new EmailBody(MESSAGE_PERSONALIZED);

        private static readonly string[] TEMPLATES = {
            @"html\ReportCard_Control.html",
            @"html\ReportCard_Neighbor.html",
            @"html\ReportCard_Lifestyle.html",
        };

        public static void Ping() { }

        public static void Send(Account account, ExcelRow entry)
        {
            try
            {
                ServicePointManager.ServerCertificateValidationCallback += (o, c, ch, er) => true;
                MailMessage mailMsg = new MailMessage();
                MailAddress fromAddress;
                if (account.DisplayName == "")
                    fromAddress = new MailAddress(account.EmailAddress);
                else
                    fromAddress = new MailAddress(account.EmailAddress, account.DisplayName);

                mailMsg.To.Add(entry.EmailAddress);
                mailMsg.From = fromAddress;
                mailMsg.IsBodyHtml = true;

                // Subject and Body
                mailMsg.Subject = String.Format("Your Monthly Energy Report Card");
                mailMsg.Body = new EmailBody(TEMPLATES[entry.ExperimentalCondition - 1]).Generate(entry);

                // Init SmtpClient and send on port 587 in my case. (Usual=port25)
                SmtpClient smtpClient = new SmtpClient(account.SmtpHostname, account.PortNumber);
                smtpClient.EnableSsl = account.SslIsEnabled;
                smtpClient.Credentials = new System.Net.NetworkCredential(account.LoginId, account.Password);

                smtpClient.Send(mailMsg);
            }
            // TODO: add exception handling for email sending
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw;
            }
        }
    }

    public class EmailBody
    {

        private static string Emphasize(string s) {
            const string EM = @"<span style='font-weight:bold;color:#ffff00;'>";
            const string ENDEM = @"</span>";
            return EM + s + ENDEM;
        }

        private static readonly string[] USER_GROUP_A = {
            "a",
            "a",
            "an",
            "an",
            "a",
            "a",
            "an",
        };

        private static readonly string[] USER_GROUP_NAMES = {
            "Neighbor",
            "Night Owl",
            "Early Bird",
            "All Nighter",
            "Midday User",
            "Sunset User",
            "Afternoon User",
        };

        private static readonly string[] USER_GROUP_TEND_TO = {
            "*UNDEFINED*",
            "use the most electricity from Midnight to Sunrise",
            "use the most electricity from Sunrise to Noon",
            "use the most electricity from Sunset to Sunrise",
            "use the most electricity from Sunrise to Sunset",
            "use the most electricity from Sunrise to Sunset",
            "use the most electricity from Sunset to Midnight",
            "use the most electricity from Noon to Midnight",
        };
    
        private static readonly string[] USER_GROUP_WHEN = {
            "*UNDEFINED*",
            "late at night",
            "in the morning",
            "all night",
            "all day",
            "in the evening",
            "in the afternoon and evening",
        };

        private static readonly string[] USER_GROUP_IMG_URLS = {
            @"https://i.ibb.co/rdCNzZ3/Neighbor.png",
            @"https://i.ibb.co/ByCLbhM/NightOwl.png",
            @"https://i.ibb.co/HxhNBXv/Early-Bird.png",
            @"https://i.ibb.co/tbdLfGd/All-Nighter.png",
            @"https://i.ibb.co/9tffwhn/Midday-Mover.jpg",
            @"https://i.ibb.co/gmg4gKG/Sunset-User.png",
            @"https://i.ibb.co/VVVBbpb/Afternoon-User.png",
        };

        private static readonly string[] RATING_TITLES = {
            @"Undefined Rating",
            Emphasize("Best!") + "Good job!",
            Emphasize("Good,") + "keep working at it!",
            Emphasize("Poor,") + "but keep working at it!",
        };

        private static readonly string[] RATING_TEXTS = {
            "Undefined rating.",
            "You used much less electricity this month than the average {0}.",
            "You used about the same amount of electricity this month as the average {0}.",
            "You used much more electricity this month than the average {0}.",
        };
    

        private enum InsertionKey
        {
            Usage,
            A,
            UserGroup,
            YouTendTo,
            When,
            Mean,
            Low,
            High,
            RatingTitle,
            RatingText,
            CursorYou,
            CursorMean,
            UserGroupImg,
        }

        private List<string> m_paragraphText = new List<string>();
        private List<InsertionKey> m_insertionKeys = new List<InsertionKey>();

        public EmailBody(string filename)
        {
            string readLine;
            List<string> allLines = new List<string>();
            string[] words;

            try
            {
                using (FileStream fs = File.Open(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (BufferedStream bs = new BufferedStream(fs))
                using (StreamReader sr = new StreamReader(bs))
                {
                    while (!(readLine = sr.ReadLine()).ToLower().Contains("<body"))
                    {
                        if (sr.EndOfStream)
                        {
                            throw new InvalidDataException("No opening <body> tag found in HTML file.");
                        }
                    }

                    while (!(readLine = sr.ReadLine()).ToLower().Contains("</body"))
                    {
                        if (sr.EndOfStream)
                        {
                            throw new InvalidDataException("No closing </body> tag found in HTML file.");
                        }

                        if (readLine != null)
                        {
                            allLines.Add(readLine);
                        }
                    }
                }

                foreach (string line in allLines)
                {
                    words = line.Trim().Split('{', '}', '\t');

                    for (int i = 0; i < words.Length; ++i)
                    {
                        switch (words[i])
                        {
                            case "USAGE":         m_insertionKeys.Add(InsertionKey.Usage);        break;
                            case "A":             m_insertionKeys.Add(InsertionKey.A);            break;
                            case "USER_GROUP":    m_insertionKeys.Add(InsertionKey.UserGroup);    break;
                            case "YOU_TEND_TO":   m_insertionKeys.Add(InsertionKey.YouTendTo);    break;
                            case "WHEN":          m_insertionKeys.Add(InsertionKey.When);         break;
                            case "MEAN":          m_insertionKeys.Add(InsertionKey.Mean);          break;
                            case "LOW":           m_insertionKeys.Add(InsertionKey.Low);          break;
                            case "HIGH":          m_insertionKeys.Add(InsertionKey.High);         break;
                            case "RATING_TITLE":  m_insertionKeys.Add(InsertionKey.RatingTitle);  break;
                            case "RATING_TEXT":   m_insertionKeys.Add(InsertionKey.RatingText);   break;
                            case "CURSOR_YOU":    m_insertionKeys.Add(InsertionKey.CursorYou);    break;
                            case "CURSOR_MEAN":   m_insertionKeys.Add(InsertionKey.CursorMean);   break;
                            case "USER_GROUP_IMG":m_insertionKeys.Add(InsertionKey.UserGroupImg); break;
                            default:
                                if (m_paragraphText.Count > m_insertionKeys.Count)
                                    m_paragraphText[m_paragraphText.Count - 1] = m_paragraphText[m_paragraphText.Count - 1] + words[i];
                                else
                                    m_paragraphText.Add(words[i]);
                                break;
                        }
                    }
                }
            }
            // TODO: add exception handling to EmailBody constructor
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw;
            }
        }

        public string Generate(ExcelRow entry)
        {
            string toReturn = "";

            for (int i = 0; i < m_paragraphText.Count; ++i)
            {
                toReturn += m_paragraphText[i];
                if (m_insertionKeys.Count > i)
                {
                    switch (m_insertionKeys[i])
                    {
                        case InsertionKey.Usage:
                            toReturn += String.Format("{0:0.#}", entry.YourEnergyUse);
                            break;                        
                        case InsertionKey.A:
                            toReturn += USER_GROUP_A[entry.Lifestyle];
                            break;
                        case InsertionKey.UserGroup:
                            toReturn += USER_GROUP_NAMES[entry.Lifestyle];
                            break;
                        case InsertionKey.YouTendTo:
                            toReturn += USER_GROUP_TEND_TO[entry.Lifestyle];
                            break;
                        case InsertionKey.When:
                            toReturn += USER_GROUP_WHEN[entry.Lifestyle];
                            break;
                        case InsertionKey.Mean:
                            toReturn += String.Format("{0:0.#}", entry.MeanEnergyUse);
                            break;
                        case InsertionKey.Low:
                            toReturn += String.Format("{0:0.#}", entry.LowestEnergyUse);
                            break;
                        case InsertionKey.High:
                            toReturn += String.Format("{0:0.#}", entry.HighestEnergyUse);
                            break;
                        case InsertionKey.RatingTitle:
                            toReturn += RATING_TITLES[entry.Rating];
                            break;
                        case InsertionKey.RatingText:
                            toReturn += String.Format(
                                RATING_TEXTS[entry.Rating],
                                Emphasize(USER_GROUP_NAMES[entry.Lifestyle])
                            );
                            break;
                        case InsertionKey.CursorYou:
                            toReturn += String.Format(
                                "{0:0.#}", 
                                (entry.YourEnergyUse - entry.LowestEnergyUse) /
                                (entry.HighestEnergyUse - entry.LowestEnergyUse) *
                                (81.4 - 1.3) + 1.3
                            );
                            break;
                        case InsertionKey.CursorMean:
                            toReturn += String.Format(
                                "{0:0.#}", 
                                (entry.MeanEnergyUse - entry.LowestEnergyUse) /
                                (entry.HighestEnergyUse - entry.LowestEnergyUse) *
                                (81.4 - 1.3) + 1.3
                            );
                            break;
                        case InsertionKey.UserGroupImg:
                            toReturn += USER_GROUP_IMG_URLS[entry.Lifestyle];
                            break;
            }
                }
            }
            return toReturn;
        }
    }
}
