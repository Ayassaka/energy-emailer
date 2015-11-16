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
        private const string MESSAGE_CONTROL = @"html\ReportCard_Control.html";
        private const string MESSAGE_GENERIC = @"html\ReportCard_Generic.html";
        private const string MESSAGE_PERSONALIZED = @"html\ReportCard_Personalized.html";


        public const int MESSAGE_TYPE_CONTROL = 0;
        public const int MESSAGE_TYPE_GENERIC = 1;
        public const int MESSAGE_TYPE_PERSONALIZED = 2;

        private static EmailBody s_emailBodyControl = new EmailBody(MESSAGE_CONTROL);
        private static EmailBody s_emailBodyGeneric = new EmailBody(MESSAGE_GENERIC);
        private static EmailBody s_emailBodyPersonalized = new EmailBody(MESSAGE_PERSONALIZED);

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
                mailMsg.Subject = String.Format("Room {0}'s weekly energy report card", entry.RoomNumber);
                switch (entry.MessageType)
                {
                    case MESSAGE_TYPE_CONTROL:
                        mailMsg.Body = s_emailBodyControl.Generate(entry.RoomNumber, entry.YourEnergyUse, entry.OtherEnergyUse, entry.BestEnergyUse, entry.Rating);
                        break;
                    case MESSAGE_TYPE_GENERIC:
                        mailMsg.Body = s_emailBodyGeneric.Generate(entry.RoomNumber, entry.YourEnergyUse, entry.OtherEnergyUse, entry.BestEnergyUse, entry.Rating);
                        break;
                    case MESSAGE_TYPE_PERSONALIZED:
                        mailMsg.Body = s_emailBodyPersonalized.Generate(entry.RoomNumber, entry.YourEnergyUse, entry.OtherEnergyUse, entry.BestEnergyUse, entry.Rating);
                        break;
                }

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
        private const int RATING_EXCELLENT = 3;
        private const int RATING_GOOD = 2;
        private const int RATING_POOR = 1;
        private const string COLOR_ACTIVE = "ff0000";
        private const string COLOR_INACTIVE = "777777";

        private enum InsertionKey
        {
            RoomNumber,
            EnergyUsedMe,
            EnergyUsedOther,
            EnergyUsedBest,
            BarGraphHeightMeTop,
            BarGraphHeightMeBottom,
            BarGraphHeightOtherTop,
            BarGraphHeightOtherBottom,
            BarGraphHeightBestTop,
            BarGraphHeightBestBottom,
            ColorExcellent,
            ColorGood,
            ColorPoor,
            imageExcellent,
            imageGood,
            imagePoor
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
                            case "YOU":
                                m_insertionKeys.Add(InsertionKey.EnergyUsedMe);
                                break;
                            case "NUM":
                                m_insertionKeys.Add(InsertionKey.RoomNumber);
                                break;
                            case "OTH":
                                m_insertionKeys.Add(InsertionKey.EnergyUsedOther);
                                break;
                            case "EFF":
                                m_insertionKeys.Add(InsertionKey.EnergyUsedBest);
                                break;
                            case "cEX":
                                m_insertionKeys.Add(InsertionKey.ColorExcellent);
                                break;
                            case "cGD":
                                m_insertionKeys.Add(InsertionKey.ColorGood);
                                break;
                            case "cPR":
                                m_insertionKeys.Add(InsertionKey.ColorPoor);
                                break;
                            case "hYOU_T":
                                m_insertionKeys.Add(InsertionKey.BarGraphHeightMeTop);
                                break;
                            case "hYOU_B":
                                m_insertionKeys.Add(InsertionKey.BarGraphHeightMeBottom);
                                break;
                            case "hOTH_T":
                                m_insertionKeys.Add(InsertionKey.BarGraphHeightOtherTop);
                                break;
                            case "hOTH_B":
                                m_insertionKeys.Add(InsertionKey.BarGraphHeightOtherBottom);
                                break;
                            case "hEFF_T":
                                m_insertionKeys.Add(InsertionKey.BarGraphHeightBestTop);
                                break;
                            case "hEFF_B":
                                m_insertionKeys.Add(InsertionKey.BarGraphHeightBestBottom);
                                break;
                            case "iEX":
                                m_insertionKeys.Add(InsertionKey.imageExcellent);
                                break;
                            case "iGD":
                                m_insertionKeys.Add(InsertionKey.imageGood);
                                break;
                            case "iPR":
                                m_insertionKeys.Add(InsertionKey.imagePoor);
                                break;
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

        public string Generate(string roomNumber, double energyUsedMe, double energyUsedOther, double energyUsedBest, int rating)
        {
            string toReturn = "";

            for (int i = 0; i < m_paragraphText.Count; ++i)
            {
                toReturn += m_paragraphText[i];
                if (m_insertionKeys.Count > i)
                {
                    switch (m_insertionKeys[i])
                    {
                        case InsertionKey.RoomNumber:
                            toReturn += roomNumber;
                            break;
                        case InsertionKey.EnergyUsedMe:
                            toReturn += String.Format("{0:0.#}", energyUsedMe);
                            break;
                        case InsertionKey.EnergyUsedOther:
                            toReturn += String.Format("{0:0.#}", energyUsedOther);
                            break;
                        case InsertionKey.EnergyUsedBest:
                            toReturn += String.Format("{0:0.#}", energyUsedBest);
                            break;
                        case InsertionKey.BarGraphHeightMeTop:
                            toReturn += BarGraphHeightTop(energyUsedMe, energyUsedOther);
                            break;
                        case InsertionKey.BarGraphHeightMeBottom:
                            toReturn += BarGraphHeightBottom(energyUsedMe, energyUsedOther);
                            break;
                        case InsertionKey.BarGraphHeightOtherTop:
                            toReturn += BarGraphHeightTop(energyUsedOther, energyUsedMe);
                            break;
                        case InsertionKey.BarGraphHeightOtherBottom:
                            toReturn += BarGraphHeightBottom(energyUsedOther, energyUsedMe);
                            break;
                        case InsertionKey.BarGraphHeightBestTop:
                            toReturn += BarGraphHeightTop(energyUsedBest, (energyUsedMe > energyUsedOther ? energyUsedMe : energyUsedOther));
                            break;
                        case InsertionKey.BarGraphHeightBestBottom:
                            toReturn += BarGraphHeightBottom(energyUsedBest, (energyUsedMe > energyUsedOther ? energyUsedMe : energyUsedOther));
                            break;
                        case InsertionKey.ColorExcellent:
                            toReturn += (rating == RATING_EXCELLENT ? COLOR_ACTIVE : COLOR_INACTIVE);
                            break;
                        case InsertionKey.ColorGood:
                            toReturn += (rating == RATING_GOOD ? COLOR_ACTIVE : COLOR_INACTIVE);
                            break;
                        case InsertionKey.ColorPoor:
                            toReturn += (rating == RATING_POOR ? COLOR_ACTIVE : COLOR_INACTIVE);
                            break;
                        case InsertionKey.imageExcellent:
                            toReturn += (rating == RATING_EXCELLENT ? @"http://s3.postimg.org/7f1yxnajj/star_active.png" : @"http://s30.postimg.org/oxs6v1u2l/star_inactive.png");
                            break;
                        case InsertionKey.imageGood:
                            toReturn += (rating == RATING_GOOD ? @"http://s3.postimg.org/7f1yxnajj/star_active.png" : @"http://s30.postimg.org/oxs6v1u2l/star_inactive.png");
                            break;
                        case InsertionKey.imagePoor:
                            toReturn += (rating == RATING_POOR ? @"http://s29.postimg.org/4l43qfsc3/frown_active.png" : @"http://s14.postimg.org/5rawx4sst/frown_inactive.png");
                            break;
                    }
                }
            }

            return toReturn;
        }

        public string BarGraphHeightTop(double thisBarUsage, double otherBarUsage)
        {
            if (thisBarUsage >= otherBarUsage || otherBarUsage <= 0)
                return "1";
            else
            {
                int height = 150 - (int)(thisBarUsage / otherBarUsage * 150.0);
                return (height >= 100 || thisBarUsage == 0 ? "100" : height.ToString());
            }
        }

        public string BarGraphHeightBottom(double thisBarUsage, double otherBarUsage)
        {
            if (thisBarUsage >= otherBarUsage || otherBarUsage <= 0)
                return "149";
            else
            {
                int height = (int)(thisBarUsage / otherBarUsage * 150.0);
                return (height <= 50 || thisBarUsage == 0 ? "50" : height.ToString());
            }
        }


    }
}
