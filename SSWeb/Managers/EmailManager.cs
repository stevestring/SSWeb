﻿using System;
using System.Net.Mail;
using System.Web;
using System.Web.Security;
using System.Net;
using SendGrid;
using SendGrid.Helpers.Mail;
using System.Configuration;
namespace SSWeb.Managers
{
    public class EmailManager
    {
        private const string EmailFrom = "noreply@360cells.com";


        public static void SendConfirmationEmail(string userName)
        {
            var user = Membership.GetUser(userName.ToString());
            var confirmationGuid = user.ProviderUserKey.ToString();
            var verifyUrl = HttpContext.Current.Request.Url.GetLeftPart
               (UriPartial.Authority) + "/Account/Verify/" + confirmationGuid;



            // Retrieve the API key from the environment variables. See the project README for more info about setting this up.
            var apiKey = ConfigurationManager.AppSettings["SendGridKey"];

            var client = new SendGridClient(apiKey);

            // Send a Single Email using the Mail Helper
            var from = new EmailAddress("noreply@360Cells.com", "360 Cells");
            var subject = "Please Verify your Account";
            var to = new EmailAddress(user.Email);

            var htmlContent = "<html><head><meta content=\"text/html; charset=utf-8\" /></head><body><p>Dear " + user.UserName +
                        ", </p><p>To verify your account, please click the following link:</p>"
                        + "<p><a href=\"" + verifyUrl + "\" target=\"_blank\">" + verifyUrl
                        + "</a></p><div>Best regards,</div><div>360cells.com</div><p>Do not forward "
                        + "this email. The verify link is private.</p></body></html>";

            var msg = MailHelper.CreateSingleEmail(from, to, subject, "", htmlContent);
            var x = client.SendEmailAsync(msg);
        }
    }
}
 
