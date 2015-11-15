using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EnergyEmailer
{
    [Serializable()]
    public class Account
    {
        public string EmailAddress
        {
            get;
            private set;
        }
        
        public string Password
        {
            get;
            private set;
        }
        
        public string SmtpHostname
        {
            get;
            private set;
        }
        
        public int PortNumber
        {
            get;
            private set;
        }
        
        public bool SslIsEnabled
        {
            get;
            private set;
        }
        
        public string LoginId
        {
            get;
            private set;
        }
        
        public string DisplayName
        {
            get;
            private set;
        }

        public Account(Account toCopy)
        {
            EmailAddress = toCopy.EmailAddress;
            LoginId = toCopy.LoginId;
            DisplayName = toCopy.DisplayName;
            SmtpHostname = toCopy.SmtpHostname;
            PortNumber = toCopy.PortNumber;
            SslIsEnabled = toCopy.SslIsEnabled;
        }

        public Account(string p_emailAddress, string p_username, string p_displayName, string p_password, string p_smtpHostname, int p_portNumber, bool p_sslIsEnabled)
        {
            EmailAddress = p_emailAddress;
            DisplayName = p_displayName;
            Password = p_password;
            SmtpHostname = p_smtpHostname;
            PortNumber = p_portNumber;
            SslIsEnabled = p_sslIsEnabled;
            LoginId = p_username;
        }
    }
}
