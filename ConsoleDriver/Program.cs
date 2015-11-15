using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;
using System.Net.Mail;
using System.Net;
using System.IO;

namespace ConsoleDriver
{
    class Program
    {
        static void Main(string[] args)
        {
            EmailBody bodyTest = new EmailBody(@"C:/Users/Richmond/Google Drive/Research/Programs/EnergyEmailer/EnergyEmailer/bin/Debug/ReportCard_Control.html");
            //Console.WriteLine(bodyTest.Generate("1111", 13.5, 5.5, 2));


            MailMessage mailMsg = new MailMessage();
            mailMsg.To.Add("rbstarbuck@gmail.com");
            mailMsg.From = new MailAddress("rbstarbuck@gmail.com"); ;
            mailMsg.IsBodyHtml = true;

            // Subject and Body
            mailMsg.Subject = "Energy use report card";
            mailMsg.Body = bodyTest.Generate("123", 4.5, 10.564, 3);

            // Init SmtpClient and send on port 587 in my case. (Usual=port25)
            SmtpClient smtpClient = new SmtpClient("smtp.gmail.com", 587);
            smtpClient.EnableSsl = true;

            System.Net.NetworkCredential credentials = new System.Net.NetworkCredential("rbstarbuck", "Poopoo8594");
            smtpClient.Credentials = credentials;

            smtpClient.Send(mailMsg);


            Console.ReadKey();
        }

        void Test()
        {
            string fileName = "C:/Users/Richmond/Google Drive/Research/Programs/EnergyEmailer/Excel Documents/first_test_93.xls";
            string connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", fileName);

            OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM [Sheet1$]", connectionString);
            DataSet ds = new DataSet();

            adapter.Fill(ds, "Test");

            DataTable data = ds.Tables["Test"];

            foreach (DataRow row in data.Rows)
            {
                Console.WriteLine("{0}\t{1}\t{2}\t{3}\t{4}\t{5}", row[0], row[1], row[2], row[3], row[4], row[5]);
                double test = Double.Parse(row[3].ToString()) + Double.Parse(row[4].ToString());
                Console.WriteLine(test);
            }

            Console.WriteLine("Done Excel");
            Console.ReadKey();


            MailMessage mailMsg = new MailMessage();
            mailMsg.To.Add("rbstarbuck@gmail.com");
            // From
            MailAddress mailAddress = new MailAddress("rbstarbuck@gmail.com");
            mailMsg.From = mailAddress;
            mailMsg.IsBodyHtml = true;

            // Subject and Body
            mailMsg.Subject = "subject";
            mailMsg.Body = "body";

            // Init SmtpClient and send on port 587 in my case. (Usual=port25)
            SmtpClient smtpClient = new SmtpClient("smtp.gmail.com", 587);
            smtpClient.EnableSsl = true;

            System.Net.NetworkCredential credentials = new System.Net.NetworkCredential("rbstarbuck", "Poopoo8594");
            smtpClient.Credentials = credentials;

            smtpClient.Send(mailMsg);


            // SHORT FORM
            //SmtpClient client = new SmtpClient("smtp.gmail.com", 587)
            //{
            //    Credentials = new NetworkCredential("rbstarbuck", "Poopoo8594"),
            //    EnableSsl = true
            //};
            
            //client.Send("rbstarbuck@gmail.com", "rbstarbuck@gmail.com", "test3", "testbody3");



            Console.WriteLine("Done email");
            Console.ReadKey();
        }
    }


    public class ExcelWorksheet
    {
        private DataTable _data;

        public int NumRows { get; private set; }

        public ExcelWorksheet(string fileName, string worksheetName)
        {
            string sourceTableName = fileName + "_" + worksheetName;
            DataSet dataSet = new DataSet();
            OleDbDataAdapter adapter = new OleDbDataAdapter
                (
                String.Format("SELECT * FROM [{0}$]", worksheetName),
                String.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", fileName)
                );

            try
            {
                adapter.Fill(dataSet, sourceTableName);
                _data = dataSet.Tables[sourceTableName];
            }
            //TODO: add exception handling to ExcelSpreadsheet constructor
            catch { throw; }

            NumRows = _data.Rows.Count;
        }

        public DataRow this[int index]
        {
            get { return _data.Rows[index]; }
        }
    }
}



public class EmailBody
{
    public const int RATING_EXCELLENT = 3;
    public const int RATING_GOOD = 2;
    public const int RATING_POOR = 1;
    public const string COLOR_ACTIVE = "ff0000";
    public const string COLOR_INACTIVE = "999999";

    public enum InsertionKey
    {
        RoomNumber,
        EnergyUsedMe,
        EnergyUsedOther,
        BarGraphHeightMe,
        BarGraphHeightOther,
        ColorExcellent,
        ColorGood,
        ColorPoor
    }

    private List<string> _paragraphText = new List<string>();
    private List<InsertionKey> _insertionKeys = new List<InsertionKey>();

    public EmailBody(string fileName)
    {
        string readLine;
        List<string> allLines = new List<string>();
        string[] words;
        
        try
        {
            using (FileStream fs = File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (BufferedStream bs = new BufferedStream(fs))
            using (StreamReader sr = new StreamReader(bs))
            {
                while (!(readLine = sr.ReadLine()).ToLower().Contains("<body"))
                {
                    if (sr.EndOfStream)
                        throw new InvalidDataException("No opening <body> tag found in HTML file.");
                }

                while (!(readLine = sr.ReadLine()).ToLower().Contains("</body"))
                {
                    if (sr.EndOfStream)
                        throw new InvalidDataException("No closing </body> tag found in HTML file.");

                    if (readLine != null)
                        allLines.Add(readLine);
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
                            _insertionKeys.Add(InsertionKey.EnergyUsedMe);
                            break;
                        case "NUM":
                            _insertionKeys.Add(InsertionKey.RoomNumber);
                            break;
                        case "OTH":
                            _insertionKeys.Add(InsertionKey.EnergyUsedOther);
                            break;
                        case "cEX":
                            _insertionKeys.Add(InsertionKey.ColorExcellent);
                            break;
                        case "cGD":
                            _insertionKeys.Add(InsertionKey.ColorGood);
                            break;
                        case "cPR":
                            _insertionKeys.Add(InsertionKey.ColorPoor);
                            break;
                        case "hYOU":
                            _insertionKeys.Add(InsertionKey.BarGraphHeightMe);
                            break;
                        case "hOTH":
                            _insertionKeys.Add(InsertionKey.BarGraphHeightOther);
                            break;
                        default:
                            if (_paragraphText.Count > _insertionKeys.Count)
                                _paragraphText[_paragraphText.Count - 1] = _paragraphText[_paragraphText.Count - 1] + words[i];
                            else
                                _paragraphText.Add(words[i]);
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

    public string Generate(string roomNumber, double energyUsedMe, double energyUsedOther, int rating)
    {
        string toReturn = "";

        for (int i = 0; i < _paragraphText.Count; ++i)
        {
            toReturn += _paragraphText[i];
            if (_insertionKeys.Count > i)
            {
                switch (_insertionKeys[i])
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
                    case InsertionKey.BarGraphHeightMe:
                        toReturn += BarGraphHeight(energyUsedMe, energyUsedOther);
                        break;
                    case InsertionKey.BarGraphHeightOther:
                        toReturn += BarGraphHeight(energyUsedOther, energyUsedMe);
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
                }
            }
        }

        return toReturn;
    }

    public string BarGraphHeight(double thisBarUsage, double otherBarUsage)
    {
        if (thisBarUsage >= otherBarUsage)
            return "1";
        else
        {
            int height = 120 - (int)(thisBarUsage / otherBarUsage * 120);
            return (height <= 90 ? height.ToString() : "90");
        }
    }
}