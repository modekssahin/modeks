using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace Modeks
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            Init_Data();
        }
        sqlsinif bgl = new sqlsinif();
        SqlCommand cmd;
        SqlDataReader dr;
        string yetkisi = "";
        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private void Init_Data()
        {
            if (Properties.Settings.Default.KullanıcıAdı != string.Empty)
            {
                if (Properties.Settings.Default.BeniHatırla == true)
                {
                    txtUsername.Text = Properties.Settings.Default.KullanıcıAdı;
                    txtPassword.Text = Properties.Settings.Default.Şifre;
                    chcRememberMe.Checked = true;
                }
                else
                {
                    txtUsername.Text = Properties.Settings.Default.KullanıcıAdı;
                    txtPassword.Text = Properties.Settings.Default.Şifre;
                }
            }
        }
        private void Save_Data()
        {
            if (chcRememberMe.Checked)
            {
                Properties.Settings.Default.KullanıcıAdı = txtUsername.Text.Trim();
                Properties.Settings.Default.Şifre = txtPassword.Text.Trim();
                Properties.Settings.Default.BeniHatırla = true;
                Properties.Settings.Default.Save();
            }
            else
            {
                Properties.Settings.Default.KullanıcıAdı = "";
                Properties.Settings.Default.Şifre = "";
                Properties.Settings.Default.BeniHatırla = false;
                Properties.Settings.Default.Save();
            }
        }

        private void giriş()
        {
            if (txtUsername.Text != "" && txtPassword.Text != "" && yetkisi != "")
            {
                string sorgu = "SELECT * FROM Giriş where kullanıcıadı=@kullanıcıadı AND şifre=@şifre AND yetki=@yetki";
                cmd = new SqlCommand(sorgu, bgl.baglanti());
                cmd.Parameters.AddWithValue("@kullanıcıadı", txtUsername.Text);
                cmd.Parameters.AddWithValue("@şifre", txtPassword.Text);
                cmd.Parameters.AddWithValue("@yetki", yetkisi);
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    if (yetkisi == "Yönetici")
                    {
                        Form2 frm = new Form2();
                        frm.kullaniciadi = txtUsername.Text;
                        frm.yetki = "Yönetici";
                        this.Hide();
                        frm.Show();
                    }
                    else if (yetkisi == "CNC")
                    {
                        Form6 frm = new Form6();
                        frm.yetki = "CNC";
                        this.Hide();
                        frm.Show();
                    }
                    else if (yetkisi == "Etiket ve Palet")
                    {
                        Form7 frm = new Form7();
                        frm.yetki = "Etiket ve Palet";
                        this.Hide();
                        frm.Show();
                    }
                    else if (yetkisi == "Membran / Press")
                    {
                        Form8 frm = new Form8();
                        frm.yetki = "Membran / Press";
                        this.Hide();
                        frm.Show();
                    }
                    else if (yetkisi == "Paket ve Kargo")
                    {
                        Form9 frm = new Form9();
                        frm.yetki = "Paket ve Kargo";
                        this.Hide();
                        frm.Show();
                    }
                    else if (yetkisi == "Sekreter")
                    {
                        Form23 frm = new Form23();
                        frm.yetki = "Sekreter";
                        frm.kullaniciadi = txtUsername.Text;
                        this.Hide();
                        frm.Show();
                    }
                    
                }
                else
                {
                    MessageBox.Show("Bilgilerinizi Kontrol Ediniz.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                bgl.baglanti().Close();
            }
            else
            {
                MessageBox.Show("Bilgilerinizi Kontrol Ediniz.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Save_Data();
            giriş();
        }

        private void txtUsername_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                giriş();
            }
        }

        private void txtPassword_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                giriş();
            }
        }

        private void txtUsername_TextChanged(object sender, EventArgs e)
        {
            if (txtUsername.Text != "" && txtPassword.Text != "")
            {
                string sorgu = "SELECT * FROM Giriş where kullanıcıadı=@kullanıcıadı AND şifre=@şifre";
                cmd = new SqlCommand(sorgu, bgl.baglanti());
                cmd.Parameters.AddWithValue("@kullanıcıadı", txtUsername.Text);
                cmd.Parameters.AddWithValue("@şifre", txtPassword.Text);
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    yetkisi = dr["yetki"].ToString();
                }
                bgl.baglanti().Close();
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void txtPassword_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void txtUsername_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
