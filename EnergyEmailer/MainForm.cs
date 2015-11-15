using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Net.Mail;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Net;

namespace EnergyEmailer
{
    public partial class MainForm : Form
    {
        private const int COL_EMAIL_ADDRESS = 1;
        private const int COL_MESSAGE_TYPE = 2;
        private const int COL_ROOM_NUMBER = 3;
        private const int COL_ENERGY_YOU = 4;
        private const int COL_ENERGY_OTHER = 5;
        private const int COL_RATING = 6;

        private bool willLoadSavedData = false;
        private BindingList<Account> accounts;

        private Excel.Application xlApp = new Excel.Application();
        private Excel.Workbooks xlWorkbooks;
        private Excel.Workbook xlWorkbook;
        private Excel.Sheets xlSheets;
        private Dictionary<string, Excel.Worksheet> xlWorksheets;
        private List<ExcelRow> writeableEntries;

        private Excel.Worksheet CurrentSheet
        {
            get
            {
                if (xlWorkbook != null)
                    return xlWorksheets[(string)comboBoxWorksheetSelector.SelectedItem];
                else
                    return null;
            }
        }
        
        public MainForm()
        {
            InitializeComponent();

            accounts = new BindingList<Account>();
            accounts.Add(new Account("<Create New Account>", "", "", "", "", 587, true));

            BinaryFormatter binaryFormatter = new BinaryFormatter();
            using (Stream stream = File.Open(@"bin\accounts.bin", FileMode.OpenOrCreate))
            {
                if (stream.Length > 0)
                {
                    try
                    {
                        List<Account> acctList = (List<Account>)binaryFormatter.Deserialize(stream);
                        foreach (Account account in acctList)
                            accounts.Add(account);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            "The stored account data in \"bin\\accounts.bin\" was corrupted. The file will be deleted, and all stored accounts will be lost.",
                            "Error");
                        stream.Close();
                        File.Delete(@"bin\accounts.bin");
                    }
                }
            }

            listboxSavedAccounts.DataSource = accounts;
            listboxSavedAccounts.DisplayMember = "EmailAddress";

            if (File.Exists(@"bin\inprog_data.bin") && File.Exists(@"bin\inprog_account.bin") && File.Exists(@"bin\inprog_count.bin"))
            {
                DialogResult loadData = MessageBox.Show(
                    "Previously loaded Excel worksheet data was found which has not yet been fully processed. Would you like to reload this data and resume email sending where it was previously left off from?\n\nNote that if you select \"No\" and load a new worksheet, this data will be lost.",
                    "Data Found",
                    MessageBoxButtons.YesNo);

                if (loadData == System.Windows.Forms.DialogResult.Yes)
                {
                    willLoadSavedData = true;
                }
            }
        }

        ~MainForm()
        {
            xlWorkbook.Close();
            xlWorkbooks.Close();
            xlApp.Quit();

            foreach (Excel.Worksheet sheet in xlWorksheets.Values)
            {
                Marshal.ReleaseComObject(sheet);
            }
            xlWorksheets = null;

            while (Marshal.ReleaseComObject(xlSheets) != 0) { }
            while (Marshal.ReleaseComObject(xlWorkbook) != 0) { }
            while (Marshal.ReleaseComObject(xlWorkbooks) != 0) { }
            while (Marshal.ReleaseComObject(xlApp) != 0) { }

            xlSheets = null;
            xlWorkbook = null;
            xlWorkbooks = null;
            xlApp = null;
        }

        private void LoadSavedData()
        {
            Account loadedAccount;
            int progCount;

            try
            {
                BinaryFormatter binaryFormatter = new BinaryFormatter();

                using (Stream dataStream = File.Open(@"bin\inprog_data.bin", FileMode.Open))
                using (Stream accountStream = File.Open(@"bin\inprog_account.bin", FileMode.Open))
                using (Stream countStream = File.Open(@"bin\inprog_count.bin", FileMode.Open))
                {
                    writeableEntries = (List<ExcelRow>)binaryFormatter.Deserialize(dataStream);
                    loadedAccount = (Account)binaryFormatter.Deserialize(accountStream);
                    progCount = (int)binaryFormatter.Deserialize(countStream);
                }

                populateFields(loadedAccount);
                labelTotalEntriesValue.Text = writeableEntries.Count.ToString();
                labelExcelWorkbookName.Text = "Previously Stored Data";

                sendAllEmails(loadedAccount, progCount);

                labelExcelWorkbookName.Text = "None";
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "An unidentified error occured. The data is likely corrupted and will be deleted.",
                    "Error");

                try
                {
                    File.Delete(@"bin\inprog_data.bin");
                }
                catch (Exception ex2) { }
                try
                {
                    File.Delete(@"bin\inprog_count.bin");
                }
                catch (Exception ex2) { }
                try
                {
                    File.Delete(@"bin\inprog_account.bin");
                }
                catch (Exception ex2) { }
            }
        }

        private void populateFields(Account accountData)
        {
            if (accountData != null)
            {
                textboxEmailAddress.Text = accountData.EmailAddress;
                textBoxUsername.Text = accountData.LoginId;
                textBoxDisplayName.Text = accountData.DisplayName;
                textboxPassword.Text = accountData.Password;
                textboxPortNumber.Text = accountData.PortNumber.ToString();
                textboxSmtpHostname.Text = accountData.SmtpHostname;
                checkboxEnableSSL.Checked = accountData.SslIsEnabled;

                if (accountData.Password == null || accountData.Password == "")
                    checkboxSavePassword.Checked = false;
                else
                    checkboxSavePassword.Checked = true;
            }
            else
            {
                textboxEmailAddress.Text = "";
                textBoxUsername.Text = "";
                textBoxDisplayName.Text = "";
                textboxPassword.Text = "";
                textboxPortNumber.Text = "587";
                textboxSmtpHostname.Text = "";
                checkboxEnableSSL.Checked = true;
                checkboxSavePassword.Checked = false;
            }
        }

        private Account newAccountFromFields()
        {
            return new Account(
                textboxEmailAddress.Text,
                textBoxUsername.Text,
                textBoxDisplayName.Text,
                textboxPassword.Text,
                textboxSmtpHostname.Text,
                Int32.Parse(textboxPortNumber.Text),
                checkboxEnableSSL.Checked
                );
        }

        private bool areFieldsValidForSending()
        {
            List<string> fields = new List<string>();
            bool isValid = true;

            if (textboxEmailAddress.Text == "")
            {
                fields.Add(" - Email address");
                isValid = false;
            }
            if (textBoxUsername.Text == "")
            {
                fields.Add(" - Username");
                isValid = false;
            }
            if (textboxPassword.Text == "")
            {
                fields.Add(" - Password");
                isValid = false;
            }
            if (textboxPortNumber.Text == "")
            {
                fields.Add(" - Port number");
                isValid = false;
            }
            if (textboxSmtpHostname.Text == "")
            {
                fields.Add(" - SMTP hostname");
                isValid = false;
            }

            if (!isValid)
            {
                MessageBox.Show("Please fill in all required fields:\n\n" + String.Join("\n", fields.ToArray()), "Error");
                return false;
            }
            else
                return true;
        }

        private bool areFieldsValidForSaving()
        {
            List<string> fields = new List<string>();
            bool isValid = true;

            if (textboxEmailAddress.Text == "")
            {
                fields.Add(" - Email address");
                isValid = false;
            }
            if (textBoxUsername.Text == "")
            {
                fields.Add(" - Username");
                isValid = false;
            }
            if (textboxPassword.Text == "" && checkboxSavePassword.Checked == true)
            {
                fields.Add(" - Password");
                isValid = false;
            }
            if (textboxPortNumber.Text == "")
            {
                fields.Add(" - Port number");
                isValid = false;
            }
            if (textboxSmtpHostname.Text == "")
            {
                fields.Add(" - SMTP hostname");
                isValid = false;
            }

            if (!isValid)
            {
                MessageBox.Show("Please fill in all required fields:\n\n" + String.Join("\n", fields.ToArray()), "Error");
                return false;
            }
            else
                return true;
        }

        private bool isWorkbookValid()
        {
            Excel.Worksheet sheet = CurrentSheet;

            if (sheet == null)
            {
                DialogResult newWorkbook = MessageBox.Show("No workbooks are currently open.\n\nWould you like to open a workbook now?.", "Error", MessageBoxButtons.YesNo);

                if (newWorkbook == System.Windows.Forms.DialogResult.Yes)
                {
                    openWorkbook();
                }

                return false;
            }

            return true;
        }

        private bool isWorksheetValid()
        {
            if (writeableEntries == null)
            {
                MessageBox.Show("Please load a worksheet.", "Error");
                return false;
            }

            return true;
        }

        private void saveAccounts()
        {
            List<Account> acctList = accounts.ToList<Account>();

            using (Stream stream = File.Open(@"bin\accounts.bin", FileMode.OpenOrCreate))
            {
                BinaryFormatter binaryFormatter = new BinaryFormatter();
                binaryFormatter.Serialize(stream, acctList.GetRange(1, acctList.Count - 1));
            }
        }

        private void openWorkbook()
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            string filepath;

            fileDialog.Filter = "Excel|*.xls;*.xlsx|All Files|*.*";
            fileDialog.Multiselect = false;

            if (fileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filepath = fileDialog.FileName;

                xlWorkbooks = xlApp.Workbooks;
                xlWorkbook = xlWorkbooks.Open(filepath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlSheets = xlWorkbook.Worksheets;
                xlWorksheets = new Dictionary<string, Excel.Worksheet>();
                foreach (Excel.Worksheet sheet in xlSheets)
                {
                    xlWorksheets.Add(sheet.Name, sheet);
                }

                comboBoxWorksheetSelector.DataSource = xlWorksheets.Keys.ToList<string>();

                labelExcelWorkbookName.Text = "\"" + Path.GetFileName(filepath) + "\"";
            }
        }

        private void openWorksheet()
        {
            if (isWorkbookValid())
            {
                writeableEntries = ExcelHandler.GetWorksheet(CurrentSheet);
                if (writeableEntries != null)
                {
                    comboBoxWorksheetSelector.Enabled = false;
                }
            }
        }

        private void closeWorksheet()
        {
            writeableEntries = null;
            try
            {
                File.Delete(@"bin\inprog_count.bin");
                File.Delete(@"bin\inprog_data.bin");
                File.Delete(@"bin\inprog_account.bin");
            }
            catch (Exception ex) { }

            comboBoxWorksheetSelector.Enabled = true;
        }

        private void sendAllEmails(Account account, int startingIndex = 0)
        {
            if (writeableEntries != null)
            {
                BinaryFormatter binaryFormatter = new BinaryFormatter();
                Dictionary<int, string> undeliverable = new Dictionary<int, string>();

                progressBarSending.Maximum = writeableEntries.Count;
                progressBarSending.Value = startingIndex;

                if (!willLoadSavedData)
                {
                    using (Stream dataStream = File.Open(@"bin\inprog_data.bin", FileMode.Create))
                    using (Stream accountStream = File.Open(@"bin\inprog_account.bin", FileMode.Create))
                    {
                        binaryFormatter.Serialize(dataStream, writeableEntries);
                        binaryFormatter.Serialize(accountStream, account);
                    }
                }

                for (int i = startingIndex; i < writeableEntries.Count; ++i)
                {
                    using (Stream stream = File.Open(@"bin\inprog_count.bin", FileMode.Create))
                    {
                        binaryFormatter.Serialize(stream, i);

                        try
                        {
                            Emailer.Send(account, writeableEntries[i]);
                        }
                        catch (Exception ex)
                        {
                            undeliverable.Add(i, writeableEntries[i].EmailAddress);
                        }

                    }

                    progressBarSending.Increment(1);
                    progressBarSending.Refresh();
                }

                if (undeliverable.Count == 0)
                {
                    MessageBox.Show("All emails sent without errors.", "Finished");
                }
                else
                {
                    string output = "Sending complete with errors. Emails to the following addresses were not sent:\n\n";
                    foreach (KeyValuePair<int, string> entry in undeliverable)
                    {
                        output += String.Format("\"{0}\": row {1}\n", entry.Value, entry.Key);
                    }
                    MessageBox.Show(output, "Finished (Errors)");
                }

                progressBarSending.Value = 0;
                closeWorksheet();
            }
        }

        private void buttonSaveAccount_Click(object sender, EventArgs e)
        {
            Account newAccount = newAccountFromFields();
            int index = listboxSavedAccounts.SelectedIndex;

            if (areFieldsValidForSaving())
            {
                if (!checkboxSavePassword.Checked)
                    newAccount = new Account(newAccount);

                if (index == 0)
                    accounts.Add(newAccount);
                else
                {
                    accounts.RemoveAt(index);
                    accounts.Insert(index, newAccount);
                    listboxSavedAccounts.SelectedIndex = index;
                }

                saveAccounts();
            }
        }

        private void listboxSavedAccounts_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listboxSavedAccounts.SelectedIndex == 0)
            {
                populateFields(null);
            }
            else
            {
                populateFields((Account)listboxSavedAccounts.SelectedValue);
            }
        }

        private void buttonDeleteAccount_Click(object sender, EventArgs e)
        {
            if (listboxSavedAccounts.SelectedIndex != 0)
            {
                accounts.Remove((Account)listboxSavedAccounts.SelectedValue);
                saveAccounts();
            }
        }

        private void buttonTestAccount_Click(object sender, EventArgs e)
        {
            if (areFieldsValidForSending())
            {
                DialogResult doTest = MessageBox.Show("This test will send an email to the address specified in the selected account from that same address. It will ensure that the credentials and settings provided are valid.\n\nAre you sure you want to run this test?", "Account Test", MessageBoxButtons.YesNo);

                if (doTest == DialogResult.Yes)
                {
                    try
                    {
                        ServicePointManager.ServerCertificateValidationCallback += (o, c, ch, er) => true;
                        Account currentAccount = newAccountFromFields();
                        MailMessage mailMsg = new MailMessage();

                        MailAddress mailAddress;
                        if (currentAccount.DisplayName == "")
                            mailAddress = new MailAddress(currentAccount.EmailAddress);
                        else
                            mailAddress = new MailAddress(currentAccount.EmailAddress, currentAccount.DisplayName);

                        mailMsg.To.Add(mailAddress);
                        mailMsg.From = mailAddress;
                        mailMsg.IsBodyHtml = true;

                        mailMsg.Subject = "Energy Report Card Emailer Account Settings Test";
                        mailMsg.Body = "This message was sent through the Energy Report Card Emailer system in order to test user-specified account settings. The account settings used in this test were valid and functional.";

                        SmtpClient smtpClient = new SmtpClient(currentAccount.SmtpHostname, (int)currentAccount.PortNumber);
                        smtpClient.EnableSsl = currentAccount.SslIsEnabled;

                        System.Net.NetworkCredential credentials = new System.Net.NetworkCredential(currentAccount.LoginId, currentAccount.Password);
                        smtpClient.Credentials = credentials;

                        smtpClient.Send(mailMsg);

                        MessageBox.Show("The test message was sent without error. Please check the email address specified in this account's settings to ensure that the test message was received.", "Test Succeeded");
                    }
                    catch (Exception exception)
                    {
                        MessageBox.Show("The account settings test failed:\n\n" + exception.Message, "Test Failed");
                    }
                }
            }
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openWorkbook();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void comboBoxWorksheetSelector_SelectedIndexChanged(object sender, EventArgs e)
        {
            labelTotalEntriesValue.Text = (CurrentSheet.UsedRange.Rows.Count - 1).ToString();
        }

        private void labelExcelWorkbookName_Click(object sender, EventArgs e)
        {
            openWorkbook();
        }

        private void buttonSend_Click(object sender, EventArgs e)
        {
            if (areFieldsValidForSending() && isWorkbookValid() && isWorksheetValid())
            {
                sendAllEmails(newAccountFromFields());
            }
        }

        private void buttonLoadWorksheet_Click(object sender, EventArgs e)
        {
            if (isWorkbookValid())
            {
                openWorksheet();
            }
        }

        private void buttonCloseWorksheet_Click(object sender, EventArgs e)
        {
            closeWorksheet();
        }

        private void MainForm_Shown(object sender, EventArgs e)
        {
            if (willLoadSavedData)
            {
                LoadSavedData();
                willLoadSavedData = false;
            }
        }

    }
}
