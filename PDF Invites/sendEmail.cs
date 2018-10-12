using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;


public abstract class Email
{
    public string subject;
    public string content;
    public string from;
    public string to;
    public string server;
    public string attachments;

    public abstract bool sendThis
    {
        get;
        set;
    }
}

public class SendThisEmail : Email
{    
    bool status = false;        
    public override bool sendThis
    {
        get 
        {            
            //server = "relay-hosting.secureserver.net";
            SmtpClient sc = new SmtpClient("exchange1.apexcom.com");
            try
            {
                _sendEmail se = new _sendEmail();
                se._sendthisEmail(sc, from, to, subject, content, attachments);
                status = true;
            }
            catch { }
            return status;
        }
        set
        {           
            status = value;
        }
    }   
 }

public class _sendEmail
{
    public void _sendthisEmail(SmtpClient server, string from, string to, string subject, string content, string attachments)
    {            

        MailAddress From = new MailAddress("jesusbautista@medicainc.com", "Jess " + (char)0xD8 + " Bautista", System.Text.Encoding.UTF8);
        MailAddress To = new MailAddress(to);
        MailMessage message = new MailMessage(From, To);
        message.Body = content;
        message.BodyEncoding = System.Text.Encoding.UTF8;
        message.Subject = subject;
        message.SubjectEncoding = System.Text.Encoding.UTF8;
        server.SendAsync(message,"Sending....");
        message.Dispose();
    }
}

