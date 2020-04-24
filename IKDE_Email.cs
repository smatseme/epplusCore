using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Configuration;

namespace Epplus_Export
{
    class IKDE_Email
    {
        string fromEmailAddress = ConfigurationManager.AppSettings["FromEmailAddress"].ToString();
        string fromEmailDisplayName = ConfigurationManager.AppSettings["FromDisplayName"].ToString();
        string[] toEmailAddress = ConfigurationManager.AppSettings["ToAddress"].Split(',');
        string[] toEmailAddressError = ConfigurationManager.AppSettings["ToAddressError"].Split(',');
        string[] CCEmailAddress = ConfigurationManager.AppSettings["CCAddress"].Split(',');
        string mailSubject = ConfigurationManager.AppSettings["Subject"].ToString();
        string mailHost = ConfigurationManager.AppSettings["SMTPHost"].ToString();
        string mailPort = ConfigurationManager.AppSettings["SMTPPort"].ToString();
        //string filename = ConfigurationManager.AppSettings["FileName"].ToString();

        public void SendNotifyMailUser(string fileName, string toEmailAddresss)
        {
            MailMessage message = new MailMessage();
            SmtpClient sc = new SmtpClient();
            try
            {
                for (int i = 0; i < toEmailAddress.Length; i++)
                {
                    message.To.Add(new MailAddress(toEmailAddress[i].ToString()));
                }

                for (int i = 0; i < CCEmailAddress.Length; i++)
                {
                    message.CC.Add(new MailAddress(CCEmailAddress[i].ToString()));
                }
                message.From = new MailAddress(fromEmailAddress, fromEmailDisplayName); 
                message.Subject = mailSubject;
                message.IsBodyHtml = false;
                message.Body = "Dear User, \n" +
                    "Please find attached Transaction Extract \n" +
                    "Use the following email for feedback: meeneshree.mohunlal@imperiallogistics.com \n" +

                    "Kind regards,";
                message.Attachments.Add(new Attachment(fileName));
                sc.Host = mailHost;
                sc.Port = Convert.ToInt32(mailPort);
                sc.UseDefaultCredentials = false;
                sc.EnableSsl = false;
                sc.Send(message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error " + ex.Message.ToString());
                throw;
            }
        }

        public void SendNotifyErrorMail(String Error, String Notify)
        {
            MailMessage message = new MailMessage();
            SmtpClient sc = new SmtpClient();
            try
            {
                for (int i = 0; i < toEmailAddressError.Length; i++)
                {
                    message.To.Add(new MailAddress(toEmailAddressError[i].ToString()));
                }
                message.From = new MailAddress(fromEmailAddress, fromEmailDisplayName); ;
                message.Subject = mailSubject;
                message.IsBodyHtml = false;
                message.Body = Notify + " " + Error;
                sc.Host = mailHost;
                sc.Port = Convert.ToInt32(mailPort);
                sc.UseDefaultCredentials = false;
                sc.EnableSsl = false;
                sc.Send(message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error " + ex.Message.ToString());
            }
        }
    }
}
