using LumiSoft.Net;
using LumiSoft.Net.IMAP.Client;
using LumiSoft.Net.Mail;
using LumiSoft.Net.Mime;
using LumiSoft.Net.POP3.Client;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsEmail
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        public string UserName
        {
            get { return this.TxtUserName.Text; }
        }
        public string Pwd
        {
            get { return this.TxtPwd.Text; }
        }
        public void functionPOP()
        {
            using (POP3_Client c = new POP3_Client())
            {
                try
                {
                    //连接POP3服务器
                    c.Connect("outlook.office365.com", 995, true);
                    //验证用户身份
                    c.Login(UserName, Pwd);  //邮件密码/smtp、pop3授权码          
                    MessageBox.Show("数量:" + c.Messages.Count.ToString());
                    if (c.Messages.Count > 0)
                    {
                        //遍历收件箱里的每一封邮件
                        var message = c.Messages[0];
                        //foreach (POP3_ClientMessage message in c.Messages)
                        //{
                        //try
                        //{
                        //mail.MarkForDeletion(); //删除邮件

                        //收件人、发件人、主题、时间等等走在mime_header里获得
                        Mail_Message mime_header = Mail_Message.ParseFromByte(message.HeaderToByte());

                        //发件人
                        if (mime_header.From != null)
                        {
                            string displayname = mime_header.From[0].DisplayName;
                            string from = mime_header.From[0].Address;
                            MessageBox.Show($"displayname:{displayname}--from{from}");
                        }

                        //收件人
                        if (mime_header.To != null)
                        {
                            StringBuilder sb = new StringBuilder();
                            foreach (Mail_t_Mailbox recipient in mime_header.To.Mailboxes)
                            {
                                string displayname = recipient.DisplayName;
                                string address = recipient.Address;
                                if (!string.IsNullOrEmpty(displayname))
                                {
                                    sb.AppendFormat("{0}({1});", displayname, address);
                                }
                                else
                                {
                                    sb.AppendFormat("{0};", address);
                                }
                            }
                        }

                        //抄送
                        if (mime_header.Cc != null)
                        {
                            StringBuilder sb = new StringBuilder();
                            foreach (Mail_t_Mailbox recipient in mime_header.Cc.Mailboxes)
                            {
                                string displayname = recipient.DisplayName;
                                string address = recipient.Address;
                                if (!string.IsNullOrEmpty(displayname))
                                {
                                    sb.AppendFormat("{0}({1});", displayname, address);
                                }
                                else
                                {
                                    sb.AppendFormat("{0};", address);
                                }
                            }
                        }

                        //发送邮件时间
                        DateTime dateTime = mime_header.Date;
                        string ContentID = mime_header.ContentID;
                        string MessageID = mime_header.MessageID;
                        string OrgMessageID = mime_header.OriginalMessageID;
                        string Subject = mime_header.Subject;

                        byte[] messageBytes = message.MessageToByte();

                        Mail_Message mime_message = Mail_Message.ParseFromByte(messageBytes);
                        if (mime_message == null)
                        {
                            //continue;
                            return;
                        }
                        string Body = mime_message.BodyText;
                        //try
                        //{
                        if (!string.IsNullOrEmpty(mime_message.BodyHtmlText))
                        {
                            //邮件内容
                            string BodyHtml = mime_message.BodyHtmlText;
                            //MessageBox.Show(BodyHtml);
                        }
                        //}
                        //catch
                        //{

                        //}
                        //}
                        //catch (Exception ex)
                        //{

                        //}
                    }
                    //}
                    //}
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
            }
        }
        public void functionIMAP()
        {

            using (IMAP_Client c = new IMAP_Client())
            {
                try
                {
                    //连接IMAP_Client服务器
                    c.Connect("outlook.office365.com", 993, true);
                    //验证用户身份
                    c.Login(UserName, Pwd);  //邮件密码/smtp、pop3授权码   

                    MessageBox.Show("数量:" + c.GetFolders(null).ToList().Count().ToString());
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }

            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            functionPOP();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            functionIMAP();
        }

    }
}
