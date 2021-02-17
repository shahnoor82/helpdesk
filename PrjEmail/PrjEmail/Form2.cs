using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;

namespace PrjEmail
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                var fromAddress = new MailAddress("helpdeskdoa@gmail.com", "DOA HelpDesk");
                var toAddress = new MailAddress("shah.noor82@gmail.com", "Ayaz Muhammad");
                const string fromPassword = "DoaHelpdesk123*";
                const string subject = "test";
                const string body = "Hey now!!";

                txtLog.Text = "1" + Environment.NewLine + txtLog.Text;
                Application.DoEvents();

                var smtp = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    Credentials = new NetworkCredential(fromAddress.Address, fromPassword),
                    Timeout = 20000
                };
                txtLog.Text = "2" + Environment.NewLine + txtLog.Text;
                Application.DoEvents();
                using (var message = new MailMessage(fromAddress, toAddress)
                {
                    Subject = subject,
                    Body = body
                })
                {
                    txtLog.Text = "3" + Environment.NewLine + txtLog.Text;
                    Application.DoEvents();
                    smtp.Send(message);
                    txtLog.Text = "Email sent: " + subject + Environment.NewLine + "to: " + toAddress + Environment.NewLine + txtLog.Text;
                }


            }
            catch (Exception ex)
            {

                txtLog.Text = ex.Message.ToArray() + txtLog.Text;
            }

           

        }
    }
}
