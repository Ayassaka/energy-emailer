using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EnergyEmailer
{
    [Serializable()]
    public class Account
    {
        private string emailAddress;
        public string EmailAddress
        {
            get { return emailAddress; }
        }

        private string password;
        public string Password
        {
            get { return password; }
        }

        private string smtpHostname;
        public string SmtpHostname
        {
            get { return smtpHostname; }
        }

        private int portNumber;
        public int PortNumber
        {
            get { return portNumber; }
        }

        private bool sslIsEnabled;
        public bool SslIsEnabled
        {
            get { return sslIsEnabled; }
        }

        private string loginId;
        public string LoginId
        {
            get { return loginId; }
        }

        private string displayName;
        public string DisplayName
        {
            get { return displayName; }
        }

        public Account(Account toCopy)
        {
            emailAddress = toCopy.emailAddress;
            loginId = toCopy.loginId;
            displayName = toCopy.displayName;
            smtpHostname = toCopy.smtpHostname;
            portNumber = toCopy.portNumber;
            sslIsEnabled = toCopy.sslIsEnabled;
        }

        public Account(string p_emailAddress, string p_username, string p_displayName, string p_password, string p_smtpHostname, int p_portNumber, bool p_sslIsEnabled)
        {
            emailAddress = p_emailAddress;
            displayName = p_displayName;
            password = p_password;
            smtpHostname = p_smtpHostname;
            portNumber = p_portNumber;
            sslIsEnabled = p_sslIsEnabled;
            loginId = p_username;
        }
    }
}
