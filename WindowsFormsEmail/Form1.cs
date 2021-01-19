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
                    MessageBox.Show("数量:"+c.Messages.Count.ToString());
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
