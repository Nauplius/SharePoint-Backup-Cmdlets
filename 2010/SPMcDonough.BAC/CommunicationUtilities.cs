#region Namespace Imports


// Standard class imports
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

// Imports that are needed for SharePoint and PowerShell
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.PowerShell;
using System.Management.Automation;

// Used for sending e-mail messages
using System.Net.Mail;

// For XML and LINQ processing
using System.Xml;
using System.Xml.Linq;

// Set resources alias
using res = SPMcDonough.BAC.Properties.Resources;


#endregion Namespace Imports


namespace SPMcDonough.BAC
{


    /// <summary>
    /// This utility class contains the methods that are used for communication purposes such as
    /// e-mail.
    /// </summary>
    internal static class CommunicationUtilities
    {


        #region Methods (Public, Static)


        /// <summary>
        /// This method is used to send an e-mail message "from the farm;" i.e., using the configured
        /// outbound e-mail settings that were specified in Central Administration. Any parameters not
        /// explicitly identified are pulled from SharePoint for purposes of preparing and sending the
        /// e-mail.
        /// </summary>
        /// <param name="toRecipients">A String that contains one or more e-mail addresses (comma-separated)
        /// to which the e-mail will be sent.</param>
        /// <param name="ccRecipients">A String containing zero or more e-mail addresses (comma-separated) to
        /// which the e-mail will be carbon-copied.</param>
        /// <param name="subject">A String containing the subject line that will be used for the e-mail</param>
        /// <param name="body">A String that contains the body of the e-mail to send.</param>
        /// <param name="isBodyHtml">TRUE if the <paramref name="body"/> parameter contains HTML, FALSE if it
        /// should be treated as plain text.</param>
        /// <remarks>If problems are encountered during the execution of this method, exceptions will
        /// propagate back to the caller. No exception handling is performed within the method.</remarks>
        public static void SendFarmEmail(String toRecipients, String ccRecipients, String subject, String body, Boolean isBodyHtml)
        {
            // We need to get each of the unspecified parameters that will be needed for sending e-mail. This
            // means drilling into the farm to get at the Central Admin web app and OutboundMailServiceInstance.
            SPWebApplication caWebApp = SPWebService.AdministrationService.WebApplications.First();

            // Setup the mail message we'll actually need
            using (MailMessage mailToSend = new MailMessage())
            {
                // Assign the sender info from the SharePoint outbound mail settings
                mailToSend.From = new MailAddress(caWebApp.OutboundMailSenderAddress);
                mailToSend.ReplyTo = new MailAddress(caWebApp.OutboundMailReplyToAddress);

                // Assign subject and body info.
                mailToSend.Subject = subject;
                mailToSend.Body = body;
                mailToSend.IsBodyHtml = isBodyHtml;

                // Assign the rest from supplied parameters as appropriate.
                mailToSend.To.Add(toRecipients);
                if (!String.IsNullOrEmpty(ccRecipients))
                {
                    mailToSend.CC.Add(ccRecipients);
                }

                // Prep the SMTP client and send.
                SmtpClient gateway = new SmtpClient(caWebApp.OutboundMailServiceInstance.Server.Address);
                gateway.Send(mailToSend);
            }
        }


        #endregion Methods (Public, Static)


    }
}
