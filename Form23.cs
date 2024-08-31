using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using MailKit.Net.Smtp;
using MimeKit;

using System.IdentityModel.Protocols.WSTrust;

namespace Modeks
{
    public partial class Form23 : Form
    {
        sqlsinif bgl = new sqlsinif();
        public string yetki;
        public string kullaniciadi;
        public string hangiformdan = "form23";
        public Form23()
        {
            InitializeComponent();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            label2.Text = DateTime.Now.ToLongDateString();
            label12.Text = DateTime.Now.ToLongTimeString();
        }

        private void Form23_Load(object sender, EventArgs e)
        {
            timer1.Start();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Form10 frm = new Form10();
            frm.kullaniciadi = kullaniciadi;
            frm.yetki = yetki;
            frm.hangiformdan = hangiformdan;
            this.Hide();
            frm.ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Form5 frm = new Form5();
            frm.kullaniciadi = kullaniciadi;
            frm.yetki = yetki;
            frm.hangiformdan = hangiformdan;
            this.Hide();
            frm.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Form3 frm = new Form3();
            frm.yetki = yetki;
            frm.kullaniciadi = kullaniciadi;
            frm.hangiformdan = hangiformdan;
            this.Hide();
            frm.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form18 frm = new Form18();
            frm.kullaniciadi = kullaniciadi;
            frm.yetki = yetki;
            frm.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form19 frm = new Form19();
            frm.kullaniciadi = kullaniciadi;
            frm.yetki = yetki;
            frm.hangiformdan = "form23";
            frm.ShowDialog();
            this.Hide();
        }

        public static void YedekAl(string veritabani)
        {
            
            string database = veritabani;
            string connectionString = @"Data Source=78.108.246.74;Initial Catalog=Modeks_2022;User ID=modeksadmin;Password=8659745Modeks;Encrypt=True;TrustServerCertificate=true;";
            string backupPath = @"C:\Yedekler";
            Guid uniqueId = Guid.NewGuid();
            string fileName = veritabani + "_" + DateTime.Now.ToString("yyyyMMdd") + "_" + uniqueId + "_backup.bak";
            string filePath = Path.Combine(backupPath, fileName);

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

               

                string backupQuery = $"BACKUP DATABASE {database} TO DISK='{filePath}'";

                try
                {
                    using (SqlCommand command = new SqlCommand(backupQuery, connection))
                    {
                        command.ExecuteNonQuery();
                        MessageBox.Show("Veritabanı yedekleme işlemi başarılı.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Yedekleme işlemi sırasında bir hata oluştu: {ex.Message}");
                }

                connection.Close();

                if (veritabani == "Modeks_2022")
                {
                    string smtpServer = "smtp.gmail.com";
                    int smtpPort = 587;
                    string smtpUser = "bscontroller2024@gmail.com";
                    string smtpPassword = "fovg swdt mkrb iisk";

                    string toEmail = "3dmodeks@gmail.com";
                    string subject = "Modeks_2022 Yedek "+ DateTime.Now.ToString("yyyyMMdd");
                    string body = "Ekte belirtilmiştir.";

                    var message = new MimeMessage();
                    message.From.Add(new MailboxAddress("BS", smtpUser));
                    message.To.Add(new MailboxAddress("Recipient", toEmail));
                    message.Subject = subject;

                    var bodyBuilder = new BodyBuilder
                    {
                        TextBody = body
                    };

                    bodyBuilder.Attachments.Add(filePath);

                    message.Body = bodyBuilder.ToMessageBody();

                    using (var client = new MailKit.Net.Smtp.SmtpClient())
                    {
                        client.Connect(smtpServer, smtpPort, MailKit.Security.SecureSocketOptions.StartTls);
                        client.Authenticate(smtpUser, smtpPassword);
                        client.Send(message);
                        client.Disconnect(true);
                    }
                }
                else
                {

                }
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            YedekAl("Modeks_2022");
            YedekAl("Modeks_Eski");
        }
   

    private void Form23_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Hide();
            Form1 form1 = System.Windows.Forms.Application.OpenForms["Form1"] as Form1;
            if (form1 != null)
            {
                form1.Show();
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
