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
                mailMsg.Subject = String.Format("Monthly Energy Report Card");
                mailMsg.Body = new EmailBody(@"html\ReportCard_UserGroup.html").Generate(entry);

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
        private static readonly string[] USER_GROUP_NAMES = {
            "Undefined User",
            "Afternoon User",
            "Early Bird",
            "Steady User",
        };

        private static readonly string[] USER_GROUP_TEND_TO = {
            "Undefined User",
            "use the most electricity from Noon to Midnight",
            "use the most electricity from Sunrise to Noon",
            "use electricity steadily all day and night",
        };

        private static readonly string[] USER_GROUP_ARE_WHO = {
            "Undefined User",
            "use the most energy in the afternoon and evening and live in similar size homes",
            "use the most energy in the morning and live in similar size homes",
            "use energy steadily all day and night, and live in similar size homes",
        };

        private static readonly string[] USER_GROUP_IMG_URLS = {
            "",
            "https://i.ibb.co/h2Fcw3S/group1.png",
            "https://i.ibb.co/ckmLXff/Picture1.png",
            "https://i.ibb.co/cJFP697/Picture2.png",
        };

        private static readonly string[] RATING_TITLES = {
            @"Undefined Rating",
            Emphasize("Good!") + "Keep working at it!",
        };

        private static readonly string[] RATING_TEXTS = {
            "Undefined rating.",
            "You used about the same amount of electricity this month as the average {0}.",
        };
    

        private enum InsertionKey
        {
            Usage,
            UserGroup,
            YouTendTo,
            PeopleWho,
            Low,
            High,
            RatingTitle,
            RatingText,
            CursorPos,
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
                            case "USER_GROUP":    m_insertionKeys.Add(InsertionKey.UserGroup);    break;
                            case "YOU_TEND_TO":   m_insertionKeys.Add(InsertionKey.YouTendTo);    break;
                            case "PEOPLE_WHO":    m_insertionKeys.Add(InsertionKey.PeopleWho);    break;
                            case "LOW":           m_insertionKeys.Add(InsertionKey.Low);          break;
                            case "HIGH":          m_insertionKeys.Add(InsertionKey.High);         break;
                            case "RATING_TITLE":  m_insertionKeys.Add(InsertionKey.RatingTitle);  break;
                            case "RATING_TEXT":   m_insertionKeys.Add(InsertionKey.RatingText);   break;
                            case "CURSOR_POS":    m_insertionKeys.Add(InsertionKey.CursorPos);    break;
                            case "USER_GROUP_IMG":m_insertionKeys.Add(InsertionKey.UserGroupImg);    break;
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
                        case InsertionKey.UserGroup:
                            toReturn += USER_GROUP_NAMES[entry.UserGroupId];
                            break;
                        case InsertionKey.YouTendTo:
                            toReturn += USER_GROUP_TEND_TO[entry.UserGroupId];
                            break;
                        case InsertionKey.PeopleWho:
                            toReturn += USER_GROUP_ARE_WHO[entry.UserGroupId];
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
                                Emphasize(USER_GROUP_NAMES[entry.UserGroupId])
                            );
                            break;
                        case InsertionKey.CursorPos:
                            toReturn += String.Format(
                                "{0:0.#}", 
                                (entry.YourEnergyUse - entry.LowestEnergyUse) /
                                (entry.HighestEnergyUse - entry.LowestEnergyUse) *
                                (85.0 - 6.0)
                            );
                            break;
                        case InsertionKey.UserGroupImg:
                            toReturn += USER_GROUP_IMG_URLS[entry.UserGroupId];
                            break;
            }
                }
            }
            return toReturn;
        }
    }
}
