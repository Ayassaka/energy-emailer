namespace EnergyEmailer
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.buttonSaveAccount = new System.Windows.Forms.Button();
            this.listboxSavedAccounts = new System.Windows.Forms.ListBox();
            this.textboxEmailAddress = new System.Windows.Forms.TextBox();
            this.labelEmailAddress = new System.Windows.Forms.Label();
            this.textboxPassword = new System.Windows.Forms.TextBox();
            this.labelPassword = new System.Windows.Forms.Label();
            this.textboxSmtpHostname = new System.Windows.Forms.TextBox();
            this.labelSmtpHost = new System.Windows.Forms.Label();
            this.textboxPortNumber = new System.Windows.Forms.TextBox();
            this.labelPortNumber = new System.Windows.Forms.Label();
            this.labelEnableSSL = new System.Windows.Forms.Label();
            this.checkboxEnableSSL = new System.Windows.Forms.CheckBox();
            this.labelSavePassword = new System.Windows.Forms.Label();
            this.checkboxSavePassword = new System.Windows.Forms.CheckBox();
            this.buttonDeleteAccount = new System.Windows.Forms.Button();
            this.buttonTestAccount = new System.Windows.Forms.Button();
            this.comboBoxWorksheetSelector = new System.Windows.Forms.ComboBox();
            this.menuStripTop = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.progressBarSending = new System.Windows.Forms.ProgressBar();
            this.labelExcelWorksheet = new System.Windows.Forms.Label();
            this.labelTotalEntriesLabel = new System.Windows.Forms.Label();
            this.labelExcelWorkbookLabel = new System.Windows.Forms.Label();
            this.labelExcelWorkbookName = new System.Windows.Forms.Label();
            this.labelTotalEntriesValue = new System.Windows.Forms.Label();
            this.buttonSend = new System.Windows.Forms.Button();
            this.buttonLoadWorksheet = new System.Windows.Forms.Button();
            this.buttonCloseWorksheet = new System.Windows.Forms.Button();
            this.labelUsername = new System.Windows.Forms.Label();
            this.textBoxUsername = new System.Windows.Forms.TextBox();
            this.textBoxDisplayName = new System.Windows.Forms.TextBox();
            this.labelAlias = new System.Windows.Forms.Label();
            this.accountFormBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.accountFormBindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.menuStripTop.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.accountFormBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.accountFormBindingSource1)).BeginInit();
            this.SuspendLayout();
            // 
            // buttonSaveAccount
            // 
            this.buttonSaveAccount.Location = new System.Drawing.Point(12, 161);
            this.buttonSaveAccount.Name = "buttonSaveAccount";
            this.buttonSaveAccount.Size = new System.Drawing.Size(94, 22);
            this.buttonSaveAccount.TabIndex = 9;
            this.buttonSaveAccount.Text = "Save Account";
            this.buttonSaveAccount.UseVisualStyleBackColor = true;
            this.buttonSaveAccount.Click += new System.EventHandler(this.buttonSaveAccount_Click);
            // 
            // listboxSavedAccounts
            // 
            this.listboxSavedAccounts.FormattingEnabled = true;
            this.listboxSavedAccounts.Location = new System.Drawing.Point(12, 33);
            this.listboxSavedAccounts.Name = "listboxSavedAccounts";
            this.listboxSavedAccounts.Size = new System.Drawing.Size(200, 121);
            this.listboxSavedAccounts.TabIndex = 0;
            this.listboxSavedAccounts.TabStop = false;
            this.listboxSavedAccounts.SelectedIndexChanged += new System.EventHandler(this.listboxSavedAccounts_SelectedIndexChanged);
            // 
            // textboxEmailAddress
            // 
            this.textboxEmailAddress.Location = new System.Drawing.Point(325, 33);
            this.textboxEmailAddress.Name = "textboxEmailAddress";
            this.textboxEmailAddress.Size = new System.Drawing.Size(158, 20);
            this.textboxEmailAddress.TabIndex = 1;
            // 
            // labelEmailAddress
            // 
            this.labelEmailAddress.AutoSize = true;
            this.labelEmailAddress.Location = new System.Drawing.Point(228, 36);
            this.labelEmailAddress.Name = "labelEmailAddress";
            this.labelEmailAddress.Size = new System.Drawing.Size(76, 13);
            this.labelEmailAddress.TabIndex = 4;
            this.labelEmailAddress.Text = "Email Address:";
            // 
            // textboxPassword
            // 
            this.textboxPassword.Location = new System.Drawing.Point(325, 106);
            this.textboxPassword.Name = "textboxPassword";
            this.textboxPassword.PasswordChar = '*';
            this.textboxPassword.Size = new System.Drawing.Size(158, 20);
            this.textboxPassword.TabIndex = 4;
            // 
            // labelPassword
            // 
            this.labelPassword.AutoSize = true;
            this.labelPassword.Location = new System.Drawing.Point(228, 109);
            this.labelPassword.Name = "labelPassword";
            this.labelPassword.Size = new System.Drawing.Size(56, 13);
            this.labelPassword.TabIndex = 6;
            this.labelPassword.Text = "Password:";
            // 
            // textboxSmtpHostname
            // 
            this.textboxSmtpHostname.Location = new System.Drawing.Point(325, 132);
            this.textboxSmtpHostname.Name = "textboxSmtpHostname";
            this.textboxSmtpHostname.Size = new System.Drawing.Size(158, 20);
            this.textboxSmtpHostname.TabIndex = 5;
            // 
            // labelSmtpHost
            // 
            this.labelSmtpHost.AutoSize = true;
            this.labelSmtpHost.Location = new System.Drawing.Point(228, 135);
            this.labelSmtpHost.Name = "labelSmtpHost";
            this.labelSmtpHost.Size = new System.Drawing.Size(91, 13);
            this.labelSmtpHost.TabIndex = 8;
            this.labelSmtpHost.Text = "SMTP Hostname:";
            // 
            // textboxPortNumber
            // 
            this.textboxPortNumber.Location = new System.Drawing.Point(325, 158);
            this.textboxPortNumber.Name = "textboxPortNumber";
            this.textboxPortNumber.Size = new System.Drawing.Size(42, 20);
            this.textboxPortNumber.TabIndex = 6;
            this.textboxPortNumber.Text = "587";
            // 
            // labelPortNumber
            // 
            this.labelPortNumber.AutoSize = true;
            this.labelPortNumber.Location = new System.Drawing.Point(228, 161);
            this.labelPortNumber.Name = "labelPortNumber";
            this.labelPortNumber.Size = new System.Drawing.Size(69, 13);
            this.labelPortNumber.TabIndex = 10;
            this.labelPortNumber.Text = "Port Number:";
            // 
            // labelEnableSSL
            // 
            this.labelEnableSSL.AutoSize = true;
            this.labelEnableSSL.Location = new System.Drawing.Point(396, 161);
            this.labelEnableSSL.Name = "labelEnableSSL";
            this.labelEnableSSL.Size = new System.Drawing.Size(66, 13);
            this.labelEnableSSL.TabIndex = 11;
            this.labelEnableSSL.Text = "Enable SSL:";
            // 
            // checkboxEnableSSL
            // 
            this.checkboxEnableSSL.AutoSize = true;
            this.checkboxEnableSSL.Checked = true;
            this.checkboxEnableSSL.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkboxEnableSSL.Location = new System.Drawing.Point(468, 161);
            this.checkboxEnableSSL.Name = "checkboxEnableSSL";
            this.checkboxEnableSSL.Size = new System.Drawing.Size(15, 14);
            this.checkboxEnableSSL.TabIndex = 7;
            this.checkboxEnableSSL.UseVisualStyleBackColor = true;
            // 
            // labelSavePassword
            // 
            this.labelSavePassword.AutoSize = true;
            this.labelSavePassword.Location = new System.Drawing.Point(228, 190);
            this.labelSavePassword.Name = "labelSavePassword";
            this.labelSavePassword.Size = new System.Drawing.Size(84, 13);
            this.labelSavePassword.TabIndex = 14;
            this.labelSavePassword.Text = "Save Password:";
            // 
            // checkboxSavePassword
            // 
            this.checkboxSavePassword.AutoSize = true;
            this.checkboxSavePassword.Location = new System.Drawing.Point(325, 190);
            this.checkboxSavePassword.Name = "checkboxSavePassword";
            this.checkboxSavePassword.Size = new System.Drawing.Size(15, 14);
            this.checkboxSavePassword.TabIndex = 8;
            this.checkboxSavePassword.UseVisualStyleBackColor = true;
            // 
            // buttonDeleteAccount
            // 
            this.buttonDeleteAccount.Location = new System.Drawing.Point(121, 161);
            this.buttonDeleteAccount.Name = "buttonDeleteAccount";
            this.buttonDeleteAccount.Size = new System.Drawing.Size(91, 22);
            this.buttonDeleteAccount.TabIndex = 10;
            this.buttonDeleteAccount.Text = "Delete Account";
            this.buttonDeleteAccount.UseVisualStyleBackColor = true;
            this.buttonDeleteAccount.Click += new System.EventHandler(this.buttonDeleteAccount_Click);
            // 
            // buttonTestAccount
            // 
            this.buttonTestAccount.Location = new System.Drawing.Point(379, 185);
            this.buttonTestAccount.Name = "buttonTestAccount";
            this.buttonTestAccount.Size = new System.Drawing.Size(104, 22);
            this.buttonTestAccount.TabIndex = 11;
            this.buttonTestAccount.Text = "Send Test Email";
            this.buttonTestAccount.UseVisualStyleBackColor = true;
            this.buttonTestAccount.Click += new System.EventHandler(this.buttonTestAccount_Click);
            // 
            // comboBoxWorksheetSelector
            // 
            this.comboBoxWorksheetSelector.FormattingEnabled = true;
            this.comboBoxWorksheetSelector.Location = new System.Drawing.Point(108, 249);
            this.comboBoxWorksheetSelector.Name = "comboBoxWorksheetSelector";
            this.comboBoxWorksheetSelector.Size = new System.Drawing.Size(154, 21);
            this.comboBoxWorksheetSelector.TabIndex = 13;
            this.comboBoxWorksheetSelector.SelectedIndexChanged += new System.EventHandler(this.comboBoxWorksheetSelector_SelectedIndexChanged);
            // 
            // menuStripTop
            // 
            this.menuStripTop.BackColor = System.Drawing.SystemColors.Control;
            this.menuStripTop.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem});
            this.menuStripTop.Location = new System.Drawing.Point(0, 0);
            this.menuStripTop.Name = "menuStripTop";
            this.menuStripTop.Size = new System.Drawing.Size(493, 24);
            this.menuStripTop.TabIndex = 20;
            this.menuStripTop.Text = "menuStripTop";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem,
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            this.openToolStripMenuItem.Size = new System.Drawing.Size(103, 22);
            this.openToolStripMenuItem.Text = "Open";
            this.openToolStripMenuItem.Click += new System.EventHandler(this.openToolStripMenuItem_Click);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(103, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // progressBarSending
            // 
            this.progressBarSending.Location = new System.Drawing.Point(10, 346);
            this.progressBarSending.Name = "progressBarSending";
            this.progressBarSending.Size = new System.Drawing.Size(471, 21);
            this.progressBarSending.TabIndex = 22;
            // 
            // labelExcelWorksheet
            // 
            this.labelExcelWorksheet.AutoSize = true;
            this.labelExcelWorksheet.Location = new System.Drawing.Point(9, 252);
            this.labelExcelWorksheet.Name = "labelExcelWorksheet";
            this.labelExcelWorksheet.Size = new System.Drawing.Size(91, 13);
            this.labelExcelWorksheet.TabIndex = 23;
            this.labelExcelWorksheet.Text = "Excel Worksheet:";
            // 
            // labelTotalEntriesLabel
            // 
            this.labelTotalEntriesLabel.AutoSize = true;
            this.labelTotalEntriesLabel.Location = new System.Drawing.Point(9, 274);
            this.labelTotalEntriesLabel.Name = "labelTotalEntriesLabel";
            this.labelTotalEntriesLabel.Size = new System.Drawing.Size(94, 13);
            this.labelTotalEntriesLabel.TabIndex = 24;
            this.labelTotalEntriesLabel.Text = "Number of Entries:";
            // 
            // labelExcelWorkbookLabel
            // 
            this.labelExcelWorkbookLabel.AutoSize = true;
            this.labelExcelWorkbookLabel.Location = new System.Drawing.Point(9, 229);
            this.labelExcelWorkbookLabel.Name = "labelExcelWorkbookLabel";
            this.labelExcelWorkbookLabel.Size = new System.Drawing.Size(89, 13);
            this.labelExcelWorkbookLabel.TabIndex = 25;
            this.labelExcelWorkbookLabel.Text = "Excel Workbook:";
            // 
            // labelExcelWorkbookName
            // 
            this.labelExcelWorkbookName.AutoSize = true;
            this.labelExcelWorkbookName.Cursor = System.Windows.Forms.Cursors.Hand;
            this.labelExcelWorkbookName.Location = new System.Drawing.Point(105, 229);
            this.labelExcelWorkbookName.Name = "labelExcelWorkbookName";
            this.labelExcelWorkbookName.Size = new System.Drawing.Size(74, 13);
            this.labelExcelWorkbookName.TabIndex = 12;
            this.labelExcelWorkbookName.Text = "None Opened";
            this.labelExcelWorkbookName.Click += new System.EventHandler(this.labelExcelWorkbookName_Click);
            // 
            // labelTotalEntriesValue
            // 
            this.labelTotalEntriesValue.AutoSize = true;
            this.labelTotalEntriesValue.Location = new System.Drawing.Point(105, 274);
            this.labelTotalEntriesValue.Name = "labelTotalEntriesValue";
            this.labelTotalEntriesValue.Size = new System.Drawing.Size(13, 13);
            this.labelTotalEntriesValue.TabIndex = 27;
            this.labelTotalEntriesValue.Text = "0";
            // 
            // buttonSend
            // 
            this.buttonSend.Location = new System.Drawing.Point(167, 298);
            this.buttonSend.Name = "buttonSend";
            this.buttonSend.Size = new System.Drawing.Size(173, 32);
            this.buttonSend.TabIndex = 16;
            this.buttonSend.Text = "Send Emails to All";
            this.buttonSend.UseVisualStyleBackColor = true;
            this.buttonSend.Click += new System.EventHandler(this.buttonSend_Click);
            // 
            // buttonLoadWorksheet
            // 
            this.buttonLoadWorksheet.Location = new System.Drawing.Point(269, 247);
            this.buttonLoadWorksheet.Name = "buttonLoadWorksheet";
            this.buttonLoadWorksheet.Size = new System.Drawing.Size(104, 22);
            this.buttonLoadWorksheet.TabIndex = 14;
            this.buttonLoadWorksheet.Text = "Load Worksheet";
            this.buttonLoadWorksheet.UseVisualStyleBackColor = true;
            this.buttonLoadWorksheet.Click += new System.EventHandler(this.buttonLoadWorksheet_Click);
            // 
            // buttonCloseWorksheet
            // 
            this.buttonCloseWorksheet.Location = new System.Drawing.Point(379, 247);
            this.buttonCloseWorksheet.Name = "buttonCloseWorksheet";
            this.buttonCloseWorksheet.Size = new System.Drawing.Size(104, 22);
            this.buttonCloseWorksheet.TabIndex = 15;
            this.buttonCloseWorksheet.Text = "Close Worksheet";
            this.buttonCloseWorksheet.UseVisualStyleBackColor = true;
            this.buttonCloseWorksheet.Click += new System.EventHandler(this.buttonCloseWorksheet_Click);
            // 
            // labelUsername
            // 
            this.labelUsername.AutoSize = true;
            this.labelUsername.Location = new System.Drawing.Point(228, 60);
            this.labelUsername.Name = "labelUsername";
            this.labelUsername.Size = new System.Drawing.Size(58, 13);
            this.labelUsername.TabIndex = 33;
            this.labelUsername.Text = "Username:";
            // 
            // textBoxUsername
            // 
            this.textBoxUsername.Location = new System.Drawing.Point(325, 57);
            this.textBoxUsername.Name = "textBoxUsername";
            this.textBoxUsername.Size = new System.Drawing.Size(158, 20);
            this.textBoxUsername.TabIndex = 2;
            // 
            // textBoxDisplayName
            // 
            this.textBoxDisplayName.Location = new System.Drawing.Point(325, 82);
            this.textBoxDisplayName.Name = "textBoxDisplayName";
            this.textBoxDisplayName.Size = new System.Drawing.Size(158, 20);
            this.textBoxDisplayName.TabIndex = 3;
            // 
            // labelAlias
            // 
            this.labelAlias.AutoSize = true;
            this.labelAlias.Location = new System.Drawing.Point(228, 85);
            this.labelAlias.Name = "labelAlias";
            this.labelAlias.Size = new System.Drawing.Size(75, 13);
            this.labelAlias.TabIndex = 36;
            this.labelAlias.Text = "Display Name:";
            // 
            // accountFormBindingSource
            // 
            this.accountFormBindingSource.DataSource = typeof(EnergyEmailer.MainForm);
            // 
            // accountFormBindingSource1
            // 
            this.accountFormBindingSource1.DataSource = typeof(EnergyEmailer.MainForm);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(493, 381);
            this.Controls.Add(this.labelAlias);
            this.Controls.Add(this.textBoxDisplayName);
            this.Controls.Add(this.textBoxUsername);
            this.Controls.Add(this.labelUsername);
            this.Controls.Add(this.buttonCloseWorksheet);
            this.Controls.Add(this.buttonLoadWorksheet);
            this.Controls.Add(this.buttonSend);
            this.Controls.Add(this.labelTotalEntriesValue);
            this.Controls.Add(this.labelExcelWorkbookName);
            this.Controls.Add(this.labelExcelWorkbookLabel);
            this.Controls.Add(this.labelTotalEntriesLabel);
            this.Controls.Add(this.labelExcelWorksheet);
            this.Controls.Add(this.progressBarSending);
            this.Controls.Add(this.comboBoxWorksheetSelector);
            this.Controls.Add(this.buttonTestAccount);
            this.Controls.Add(this.buttonDeleteAccount);
            this.Controls.Add(this.checkboxSavePassword);
            this.Controls.Add(this.labelSavePassword);
            this.Controls.Add(this.checkboxEnableSSL);
            this.Controls.Add(this.labelEnableSSL);
            this.Controls.Add(this.labelPortNumber);
            this.Controls.Add(this.textboxPortNumber);
            this.Controls.Add(this.labelSmtpHost);
            this.Controls.Add(this.textboxSmtpHostname);
            this.Controls.Add(this.labelPassword);
            this.Controls.Add(this.textboxPassword);
            this.Controls.Add(this.labelEmailAddress);
            this.Controls.Add(this.textboxEmailAddress);
            this.Controls.Add(this.listboxSavedAccounts);
            this.Controls.Add(this.buttonSaveAccount);
            this.Controls.Add(this.menuStripTop);
            this.MainMenuStrip = this.menuStripTop;
            this.Name = "MainForm";
            this.Text = "Energy Report Card Emailer";
            this.Shown += new System.EventHandler(this.MainForm_Shown);
            this.menuStripTop.ResumeLayout(false);
            this.menuStripTop.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.accountFormBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.accountFormBindingSource1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonSaveAccount;
        private System.Windows.Forms.ListBox listboxSavedAccounts;
        private System.Windows.Forms.TextBox textboxEmailAddress;
        private System.Windows.Forms.Label labelEmailAddress;
        private System.Windows.Forms.TextBox textboxPassword;
        private System.Windows.Forms.Label labelPassword;
        private System.Windows.Forms.TextBox textboxSmtpHostname;
        private System.Windows.Forms.Label labelSmtpHost;
        private System.Windows.Forms.TextBox textboxPortNumber;
        private System.Windows.Forms.Label labelPortNumber;
        private System.Windows.Forms.Label labelEnableSSL;
        private System.Windows.Forms.CheckBox checkboxEnableSSL;
        private System.Windows.Forms.Label labelSavePassword;
        private System.Windows.Forms.CheckBox checkboxSavePassword;
        private System.Windows.Forms.Button buttonDeleteAccount;
        private System.Windows.Forms.BindingSource accountFormBindingSource;
        private System.Windows.Forms.BindingSource accountFormBindingSource1;
        private System.Windows.Forms.Button buttonTestAccount;
        private System.Windows.Forms.ComboBox comboBoxWorksheetSelector;
        private System.Windows.Forms.MenuStrip menuStripTop;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ProgressBar progressBarSending;
        private System.Windows.Forms.Label labelExcelWorksheet;
        private System.Windows.Forms.Label labelTotalEntriesLabel;
        private System.Windows.Forms.Label labelExcelWorkbookLabel;
        private System.Windows.Forms.Label labelExcelWorkbookName;
        private System.Windows.Forms.Label labelTotalEntriesValue;
        private System.Windows.Forms.Button buttonSend;
        private System.Windows.Forms.Button buttonLoadWorksheet;
        private System.Windows.Forms.Button buttonCloseWorksheet;
        private System.Windows.Forms.Label labelUsername;
        private System.Windows.Forms.TextBox textBoxUsername;
        private System.Windows.Forms.TextBox textBoxDisplayName;
        private System.Windows.Forms.Label labelAlias;
    }
}