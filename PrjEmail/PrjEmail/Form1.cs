using System;
using System.Configuration;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Web;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Net.Mail;

namespace PrjEmail
{
    public partial class Form1 : Form
    {

        protected int ReadInputStreamCharByChar = 0;
        
        protected enum service_state { STARTED, PAUSED, STOPPED };
        protected service_state state = service_state.STARTED;

        public static DateTime heartbeat_datetime = DateTime.Now;


        protected string SubjectMustContain;
        protected string SubjectCannotContain;
        protected string[] SubjectCannotContainStrings;

        protected string FromMustContain;
        protected string FromCannotContain;
        protected string[] FromCannotContainStrings;

        protected string TrackingIdString;

        protected string conStr = ""; //"server=173.212.205.20;uid=test;pwd=test;Database=HelpDesk;Connect Timeout=600;MultipleActiveResultSets=True;";
        //protected string conStr = "server=.;uid=sa;pwd=abc.123;Database=HelpDesk;Connect Timeout=600;MultipleActiveResultSets=True;";

        int pop3port;
        bool pop3use_ssl;
        string pop3server;
        string pop3emailid;
        string pop3emailpwd;

        int smtpport;
        bool smtpuse_ssl;
        string smtpserver;
        string smtpemailid;
        string smtpemailpwd;

        bool vStarStatus;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            receive_emails();
        }

        private void receive_emails()
        {
            //int port = 110;
            //bool use_ssl = false;
            //string pop3server = "mail.cpbdc.com.pk";
            //string emailid = "ayaz@cpbdc.com.pk";
            //string emailpwd = "bitline2020";

            System.Net.ServicePointManager.Expect100Continue = false;

            string[] messages = null;
            Regex regex = new Regex("\r\n");
            string[] test_message_text = new string[100];
            POP3Client.POP3client client = null;

            try
            {
                client = new POP3Client.POP3client(ReadInputStreamCharByChar);

                txtLog.Text = "****connecting to server:" + Environment.NewLine;

                txtLog.Text += client.connect(pop3server, pop3port, pop3use_ssl) + Environment.NewLine;

                txtLog.Text += "sending POP3 command USER" + Environment.NewLine;
                txtLog.Text += client.USER(pop3emailid) + Environment.NewLine;

                txtLog.Text += "sending POP3 command PASS" + Environment.NewLine;
                txtLog.Text += client.PASS(pop3emailpwd) + Environment.NewLine;

                txtLog.Text += "sending POP3 command STAT" + Environment.NewLine;
                txtLog.Text += client.STAT() + Environment.NewLine;

                txtLog.Text += "sending POP3 command LIST" + Environment.NewLine;
                string list;
                list = client.LIST();
                txtLog.Text += "list follows:" + Environment.NewLine;
                txtLog.Text += list + Environment.NewLine;

                messages = regex.Split(list);

                //txtLog.Text += client.RETR(Convert.ToInt16(txtMsgNo.Text));
                //txtLog.Text += client.QUIT();
                //return;

                //--------email fetching - header,body etc
                string message;
                int message_number = 0;
                int start;
                int end;


                start = 1;
                end = messages.Length - 1;



                // loop through the messages
                for (int i = start; i < end; i++)
                {
                    heartbeat_datetime = DateTime.Now; // because the watchdog is watching

                    if (state != service_state.STARTED || vStarStatus==false)
                    {
                        break;
                    }

                    // fetch the message

                    //txtLog.Text += "i: " + Convert.ToString(i);

                    int space_pos = messages[i].IndexOf(" ");
                    message_number = Convert.ToInt32(messages[i].Substring(0, space_pos));
                    message = client.RETR(message_number);



                    // break the message up into lines
                    string[] lines = regex.Split(message);
                    string fromEmail = "";
                    string from = "";
                    string subject = "";
                    bool vBodyStart = false;
                    string vBody = "";
                    bool encountered_subject = false;
                    bool encountered_from = false;


                    // Loop through the lines of a message.
                    // Pick out the subject and body
                    for (int j = 0; j < lines.Length; j++)
                    {

                        if (state != service_state.STARTED || vStarStatus == false)
                        {
                            break;
                        }

                        // We know from
                        // http://www.devnewsgroups.net/group/microsoft.public.dotnet.framework/topic62515.aspx
                        // that headers can be lowercase too.

                        if ((lines[j].IndexOf("Subject: ") == 0 || lines[j].IndexOf("subject: ") == 0)
                        && !encountered_subject)
                        {
                            subject = lines[j].Replace("Subject: ", "");
                            subject = subject.Replace("subject: ", ""); // try lowercase too
                            subject += maybe_append_next_line(lines, j);

                            encountered_subject = true;
                        }
                        else if (lines[j].IndexOf("Return-Path: <") == 0 && fromEmail == "")
                        {
                            fromEmail = lines[j].Replace("Return-Path: <", "").Replace(">", "");
                            fromEmail += maybe_append_next_line(lines, j);

                        }
                        else if (lines[j].IndexOf("From: ") == 0 && !encountered_from)
                        {
                            from = lines[j].Replace("From: ", "");
                            if (from.IndexOf("<")>=0)
                                from = from.Substring(0, from.IndexOf("<")).Trim();
                            encountered_from = true;
                            from += maybe_append_next_line(lines, j);

                        }

                        else if (lines[j].IndexOf("from: ") == 0 && !encountered_from)
                        {
                            from = lines[j].Replace("From: ", "");
                            if (from.IndexOf("<") >= 0)
                                from = from.Substring(0, from.IndexOf("<")).Trim();
                            encountered_from = true;
                            from += maybe_append_next_line(lines, j);
                        }
                        else if ((lines[j].IndexOf("--00000") >= 0 || lines[j].IndexOf("X-Spam-Flag:") >= 0 || lines[j].IndexOf("--_000_") >= 0 || lines[j].IndexOf("\t00000000-") >= 0) && !vBodyStart)
                        {
                            vBodyStart = true;
                        }
                        else if (vBodyStart == true && (lines[j].IndexOf("--00000") >= 0 || lines[j].IndexOf("[]") >= 0 || lines[j].IndexOf("--_000_") >= 0))
                        {
                            vBodyStart = false;
                        }
                        else if (vBodyStart == true && (lines[j].IndexOf("Content-Type: text/plain") >= 0 || lines[j].IndexOf("Content-Transfer-Encoding:") >= 0 || lines[j].IndexOf("X-MS-Exchange-Transport-CrossTenantHeadersStamped:") >= 0))
                        {
                            vBodyStart = true;
                        }
                        else if (vBodyStart == true)
                        {
                            vBody += lines[j] + Environment.NewLine;
                        }





                    } // end for each line
                    txtLog.Text = "\n Email # : " + message_number.ToString() + Environment.NewLine;
                    txtLog.Text += "From Email: " + fromEmail + Environment.NewLine;
                    txtLog.Text += "From: " + from + Environment.NewLine;
                    txtLog.Text += "Subject: " + subject + Environment.NewLine;
                    txtLog.Text += "Body: " + vBody + Environment.NewLine;

                    if (pop3emailid != fromEmail && fromEmail!="")
                    {
                        Int64 vTicketNo = insert_complaint(fromEmail, from, subject, vBody);
                    }
                    //this.SendEmail(vTicketNo.ToString(), subject,"email body", "ayaz@cpbdc.com.pk", fromEmail);

                    txtLog.Text += client.DELE(message_number);

                    System.Threading.Thread.Sleep(1000);
                    Application.DoEvents();

                }  // end for each message




                //txtLog.Text += client.RETR(Convert.ToInt32(txtMsgNo.Text));

                txtLog.Text += client.QUIT() + Environment.NewLine;

                client = null;


                ////----now sending email from outbox
                txtLog.Text += "Outbox Connected." + Environment.NewLine;
                fetch_outbox();
                txtLog.Text += "Outbox disconnected." + Environment.NewLine;
                //-------------------


                timer1.Enabled = true;
                timer1.Start();
            }
            catch (Exception exc)
            {
                txtLog.Text += "Exception trying to talk to pop3 server";
                txtLog.Text += exc;
                timer1.Enabled = true;
                timer1.Start();
                return;
            }

        
        }


        private void SendEmail(string ticketNo, string subject, string emailBody, string toEmail)
        {
            //using (MailMessage mm = new MailMessage(smtpemailid, toEmail))
            //{
            //    ticketNo = string.Concat("000000", ticketNo);
            //    ticketNo = ticketNo.Substring(ticketNo.Length - 6, 6);
            //    mm.Subject = "[Ticket # "+ ticketNo +"] " + subject;
            //    mm.Body = emailBody; 

            //    //foreach (string filePath in openFileDialog1.FileNames)
            //    //{
            //    //    if (File.Exists(filePath))
            //    //    {
            //    //        string fileName = Path.GetFileName(filePath);
            //    //        mm.Attachments.Add(new Attachment(filePath));
            //    //    }
            //    //}
            //    mm.IsBodyHtml = false;
            //    SmtpClient smtp = new SmtpClient();
            //    smtp.Host = smtpserver;  //"smtp.gmail.com";                
            //    smtp.EnableSsl = smtpuse_ssl;
            //    NetworkCredential NetworkCred = new NetworkCredential(smtpemailid, smtpemailpwd);
            //    smtp.UseDefaultCredentials = false;
            //    smtp.Credentials = NetworkCred;
            //    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            //    smtp.Port = smtpport; // 587;
            //    smtp.Send(mm);
            //    txtLog.Text = "Email sent:  [Ticket # " + ticketNo + "] " + subject + Environment.NewLine + Environment.NewLine + txtLog.Text;
            //}


            /******************************************************************************/
            
            ticketNo = string.Concat("000000", ticketNo);
            ticketNo = ticketNo.Substring(ticketNo.Length - 6, 6);
            subject = "[Ticket # " + ticketNo + "] " + subject;
            

            var smtp = new SmtpClient
            {
                Host = smtpserver,
                Port = smtpport,
                EnableSsl = smtpuse_ssl,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                Credentials = new NetworkCredential(smtpemailid, smtpemailpwd),
                Timeout = 20000
            };
            using (var message = new MailMessage(smtpemailid, toEmail)
            {
                Subject = subject,
                Body = emailBody
            })
            {
                smtp.Send(message);
                txtLog.Text = "Email sent:  [Ticket # " + ticketNo + "] " + subject + Environment.NewLine + Environment.NewLine + txtLog.Text;
            }




        }

        ///////////////////////////////////////////////////////////////////
        public void start()
        {

            // call do_work()
            txtLog.Text += "starting";
            state = service_state.STARTED;
        }

        ///////////////////////////////////////////////////////////////////
        public void pause()
        {
            txtLog.Text += "pausing";
            state = service_state.PAUSED;
        }

        ///////////////////////////////////////////////////////////////////
        public void stop()
        {
            txtLog.Text += "stopping";
            state = service_state.STOPPED;
        }


        ///////////////////////////////////////////////////////////////////////
        protected string maybe_append_next_line(string[] lines, int j)
        {
            string s = "";
            if (j + 1 < lines.Length)
            {
                int pos = -1;

                // find first non space, non tab
                for (int i = 0; i < lines[j + 1].Length; i++)
                {
                    String c = lines[j + 1].Substring(i, 1);
                    if (c == "\t" || c == " ")
                    {
                        continue;
                    }
                    else
                    {
                        pos = i;
                        break;
                    }
                }

                // this line is part of the previous header, so return it
                if (pos > 0)
                {
                    s = " ";
                    s = lines[j + 1].Substring(pos);
                }
            }
            return s;
        }

        ///////////////////////////////////////////////////////////////////////
        protected Int64 insert_complaint(string vemail, string vname, string vsubject, string vbody)
        {
            try
            {
                Int64 vcid = getMaxID();
                int vforwardTo = getForwardTo(vemail);

                SqlConnection cn2 = new SqlConnection();
                cn2.ConnectionString = conStr;
                Cursor.Current = Cursors.WaitCursor;
                cn2.Open();
                string vSql = string.Format(@"INSERT INTO Complaints (ComplaintID, ComplaintDate, EntryDateTime, complainantName, complainantFatherName, complainantCNIC, SMSMobileNo, ContactNo, Email, City, Addess1, Address2, ComplaintTypeID, ComplaintNatureID, PriorityID, AssetID, Subject, ComplaintDescription, ForwardTo, UserID) 
                                              VALUES ({0}, cast(CONVERT(varchar, GETDATE(), 101) as smalldatetime), GETDATE(), '{1}', '-', '-', '-', '-', '{2}', '-', '-', '-', 1, 1, 1, null, '{3}', '{4}', {5}, 1)", vcid, vname, vemail, vsubject, vbody.Replace("'","''"), vforwardTo.ToString());

                SqlCommand ucmd = new SqlCommand(vSql, cn2);
                int vRcd = ucmd.ExecuteNonQuery();
                ucmd.Dispose();
                cn2.Close();
                cn2.Dispose();

                return vcid;
            }
            catch (Exception exc)
            {
                
                throw exc;
            }
        }


        private void fetch_outbox()
        {
            try
            {
                
               
                SqlConnection cn = new SqlConnection();
                cn.ConnectionString = conStr;
                Cursor.Current = Cursors.WaitCursor;
                cn.Open();

                if (cn.State == ConnectionState.Open)
                {

                    //fetching the record
                    SqlCommand scmd = new SqlCommand("Select ComplaintID, ToEmail, Subject, EmailBody, IsSent, IsFailed from Outbox Where IsSent=0 Order By ComplaintID", cn);
                    SqlDataReader sdr = scmd.ExecuteReader();

                    if (sdr.HasRows)
                    {

                        while (sdr.Read())
                        {
                            if (vStarStatus == false)
                            {
                                break;
                            }
                            if (sdr["ToEmail"].ToString().Length > 0)
                            {
                                this.SendEmail(sdr["ComplaintID"].ToString(), sdr["Subject"].ToString(), sdr["EmailBody"].ToString(), sdr["ToEmail"].ToString());

                                //mark sent if successfully sent
                                
                                SqlCommand ucmd = new SqlCommand("Update Outbox Set IsSent=1, IsFailed=0 Where ComplaintID=" + sdr["ComplaintID"].ToString(), cn);
                                int vRcd = ucmd.ExecuteNonQuery();
                                ucmd.Dispose();

                                // display log info
                                //txtLog.Text = "To#: " + vToNo + ".   (Send Successfully)\r\n" + " Message: " + vSMSBody + "\r\n" + "\r\n" + TxtResult.Text;

                            }
                            else
                            {
                                SqlCommand ucmd1 = new SqlCommand("Update Outbox Set IsSent=1, IsFailed=1 Where ComplaintID=" + sdr["ComplaintID"].ToString(), cn);
                                int vRcd = ucmd1.ExecuteNonQuery();
                                ucmd1.Dispose();
                            }

                            System.Threading.Thread.Sleep(1000);
                            Application.DoEvents();

                        } //while




                    }
                    scmd.Dispose();
                    sdr.Close();
                    sdr.Dispose();
                    cn.Close();
                    cn.Dispose();
                }

            }
            catch (Exception ex)
            {
                
               txtLog.Text = "Fetching outbox Error: " + ex.Message;
            }

        }


        public Int64 getMaxID()
        {
            try
            {
                

                string query = string.Format(@"SELECT COALESCE (MAX(ComplaintID), 0)+1 AS ComplaintID FROM Complaints");
                DataSet ds = ExecuteForDataSet(query);
                Int64 vID = Convert.ToInt64(ds.Tables[0].Rows[0][0]);
                ds.Dispose();
                

                return vID;
            }
            catch (Exception exc)
            {

                throw exc;
            }
        }


        public int getForwardTo(string vFromEmail)
        {
            try
            {
                int vID = 0;
                string query = "";
                ///1st cehck assigned agent against from email id
                query = string.Format(@"SELECT TOP 1 AgentID FROM AgentsSetting Where FromEmail='{0}'", vFromEmail);
                DataSet ds = ExecuteForDataSet(query);
                if (ds.Tables[0].Rows.Count > 0)
                    vID = Convert.ToInt32(ds.Tables[0].Rows[0][0]);
                else
                    vID = 0;
                ds.Dispose();

                ///2nd get  default agent against if from email id doesnt have any agent assigned
                if (vID == 0)
                {
                    query = string.Format(@"SELECT TOP 1 AgentID FROM AgentsSetting Where IsDefault=1");
                    ds = ExecuteForDataSet(query);
                    if (ds.Tables[0].Rows.Count > 0)
                        vID = Convert.ToInt32(ds.Tables[0].Rows[0][0]);
                    else
                        vID = 0;

                    ds.Dispose();
                }

                if (vID == 0)
                    vID = 1;


                return vID;
            }
            catch (Exception exc)
            {

                throw exc;
            }
        }


        public DataSet ExecuteForDataSet(string strQuery)
        {
            SqlConnection con = new SqlConnection();
            SqlDataAdapter myAdapter = new SqlDataAdapter();
            DataSet ds = new DataSet();
            SqlCommand cmd;
            con.ConnectionString = conStr;
            try
            {
                con.Open();
                cmd = new SqlCommand(strQuery);
                cmd.Connection = con;
                cmd.CommandType = CommandType.Text;
                myAdapter.SelectCommand = cmd;
                cmd.CommandTimeout = 1000;
                myAdapter.Fill(ds);
                con.Close();
                con.Dispose();

                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                con.Close();
            }
        }

        public DataTable getEmailSettings()
        {
            string vQuery = "select TOP 1 * From EmailSettings";
            DataTable dt = new DataTable();
            SqlConnection Conn = new SqlConnection();
            Conn = new SqlConnection(conStr);
            SqlCommand QueryCmd = new SqlCommand(vQuery, Conn);
            try
            {
                Conn.Open();
                QueryCmd.ExecuteNonQuery();
                SqlDataAdapter adpater = new SqlDataAdapter(QueryCmd);
                adpater.Fill(dt);
            }
            catch (Exception ex)
            { throw ex; }
            finally
            {
                Conn.Close();
            }
            return dt;
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Stop();
            timer1.Enabled = false;
            receive_emails();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            try
            {
                conStr = "server=" + PrjEmail.Properties.Settings.Default.DBServer + ";uid=" + PrjEmail.Properties.Settings.Default.DBUser + ";pwd=" + PrjEmail.Properties.Settings.Default.DBPassword + ";Database=" + PrjEmail.Properties.Settings.Default.DBName + ";Connect Timeout=10000;MultipleActiveResultSets=True;";

                lblDBName.Text = "DB Server: " + PrjEmail.Properties.Settings.Default.DBServer;
                DataTable dtES = getEmailSettings();
                if (dtES.Rows.Count > 0)
                {
                    pop3emailid = dtES.Rows[0]["POPEmail"].ToString();
                    pop3emailpwd = dtES.Rows[0]["POPPassword"].ToString();
                    pop3server = dtES.Rows[0]["POPServer"].ToString();
                    pop3port = int.Parse(dtES.Rows[0]["POPPort"].ToString());
                    pop3use_ssl = Convert.ToBoolean(dtES.Rows[0]["POPSSL"].ToString());

                    smtpemailid = dtES.Rows[0]["SMTPEmail"].ToString();
                    smtpemailpwd = dtES.Rows[0]["SMTPPassword"].ToString();
                    smtpserver = dtES.Rows[0]["SMTPServer"].ToString();
                    smtpport = int.Parse(dtES.Rows[0]["SMTPPort"].ToString());
                    smtpuse_ssl = Convert.ToBoolean(dtES.Rows[0]["SMTPSSL"].ToString());

                    lblDBName.Text += " :  Connected";
                }

                
                //MessageBox.Show(conStr);
            }
            catch (Exception ex)
            {
                throw ex;
            }           

        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            vStarStatus = true;
            timer1.Start();
            timer1.Enabled = true;
            btnStop.Enabled = true;
            btnStart.Enabled = false;
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            vStarStatus = false;
            timer1.Stop();
            timer1.Enabled = false;
            btnStop.Enabled = false;
            btnStart.Enabled = true;

        }

        

    }
}
