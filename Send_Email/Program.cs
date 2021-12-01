using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;


namespace Send_Email
{
    public class CommunicationUtilityService
    {
        /*
         - Pass "null" for Mail_CC,Mail_BCC and Mail_Attachment if these are not neede
        */
        public string EmailSender(string Mail_From, string[] Mail_To, string[] Mail_CC, string[] Mail_BCC, string Mail_Subject, string Mail_body,
            string[] Mail_Attachment, string Mail_Display_Name, string Mail_Host, bool Mail_Ssl, string Mail_Password, int Mail_Port)
        {
            try
            {
                MailMessage mail = new MailMessage();
                mail.From = new MailAddress(Mail_From);
                //Create multiple email address for to
                for (int i = 0; i < Mail_To.Length; i++)
                {
                    mail.To.Add(new MailAddress(Mail_To[i]));
                }
                //Create multiple email address for CC
                if (Mail_CC != null)
                {
                    for (int i = 0; i < Mail_CC.Length; i++)
                    {
                        mail.CC.Add(new MailAddress(Mail_CC[i]));
                    }
                }
                //Create multiple email address for BCC
                if (Mail_BCC != null)
                {
                    for (int i = 0; i < Mail_BCC.Length; i++)
                    {
                        mail.Bcc.Add(new MailAddress(Mail_BCC[i]));
                    }
                }
                //Email Attachment
                byte[] file = null;
                if (Mail_Attachment != null)
                {
                    for (int i = 0; i < Mail_Attachment.Length; i++)
                    {
                        Attachment MailAttachment = new Attachment(Mail_Attachment[i], MediaTypeNames.Application.Pdf);
                        ContentDisposition disposition = MailAttachment.ContentDisposition;
                        disposition.CreationDate = File.GetCreationTime(Mail_Attachment[i]);
                        disposition.ModificationDate = File.GetLastWriteTime(Mail_Attachment[i]);
                        disposition.ReadDate = File.GetLastAccessTime(Mail_Attachment[i]);
                        disposition.FileName = Path.GetFileName(Mail_Attachment[i]);
                        disposition.Size = new FileInfo(Mail_Attachment[i]).Length;
                        disposition.DispositionType = DispositionTypeNames.Attachment;

                        mail.Attachments.Add(MailAttachment);
                        using (var stream = new FileStream(Mail_Attachment[i], FileMode.Open, FileAccess.Read))
                        {
                            using (var reader = new BinaryReader(stream))
                            {
                                string fileName = Path.GetFileName(Mail_Attachment[i]).ToString();
                                file = reader.ReadBytes((int)stream.Length);
                                file = null;
                            }
                        }
                    }
                }
                mail.Subject = Mail_Subject;
                mail.Body = Mail_body;
                mail.BodyEncoding = Encoding.UTF8;
                mail.IsBodyHtml = true;

                SmtpClient smtp = new SmtpClient();
                smtp.Host = Mail_Host;
                smtp.EnableSsl = Mail_Ssl;
                smtp.UseDefaultCredentials = true;
                smtp.Credentials = new NetworkCredential(Mail_From, Mail_Password);
                smtp.Port = Mail_Port;



                smtp.Send(mail);
            }
            catch (Exception exc)
            {
                return "Fail";

            }
            return "Success";
        }

    }


    class Program
    {
        static void Main(string[] args)
        {
            string[] toemail = { "swapan.acharya@itpluspoint.com" };
            string[] ccemail = {"hiteshmalik.itpp@outlook.com" };
            string[] bccemail = { "upasana.majhi.itpp@outlook.com" };
            string[] Mail_Attachment = { @"C:\ITPlusPoint\Category Data.docx" , @"C:\ITPlusPoint\Parts3.jpg" };


            CommunicationUtilityService oCommunicationUtilityService = new CommunicationUtilityService();
            oCommunicationUtilityService.EmailSender("khitesh932@gmail.com", toemail,ccemail,bccemail,"Testing for functionality 3","Testing email for multiple cc,BCC & Attachment", Mail_Attachment, "Hitesh", "smtp.gmail.com",true, "",587);


        }
    }
}
