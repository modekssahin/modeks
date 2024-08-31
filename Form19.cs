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
    public partial class Form19 : Form
    {
        public Form19()
        {
            InitializeComponent();
        }
        sqlsinif bgl = new sqlsinif();
        public string yetki;
        public string kullaniciadi;
        public string hangiformdan;

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
        string tarih = Convert.ToDateTime(DateTime.Now).ToString("yyyy-MM-dd HH:mm:ss");
        string izin;

        private void Hesapla()
        {
            izin = "";

            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT * FROM Siparişler where SiparisNo=@SiparisNo";
            komut.Parameters.AddWithValue("@SiparisNo", textBox2.Text);
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;
            SqlDataReader dr;
            dr = komut.ExecuteReader();
            if (dr.Read())
            {

                label46.Text = dr["SiparişTarihi"].ToString();
                label1.Text = dr["TeslimTarihi"].ToString();
                label5.Text = dr["OnayTarihi"].ToString();
                label7.Text = dr["Müşteri"].ToString();
                label4.Text = dr["PaketTarihi"].ToString();
                label11.Text = dr["PaketSayısı"].ToString();

                if (dr["TeslimEdilenTarih"].ToString().Length < 3)
                {
                    string tarih2 = Convert.ToDateTime(dr["TeslimTarihi"]).ToString("yyyy-MM-dd HH:mm:ss");

                    DateTime tarihObj = DateTime.ParseExact(tarih, "yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                    DateTime labelTarihObj = DateTime.ParseExact(tarih2, "yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);

                    int gunFarki = (labelTarihObj - tarihObj).Days;

                    label1.Text = dr["TeslimTarihi"].ToString() + "  " + gunFarki.ToString() + " gün var.";

                }
                else if (dr["TeslimEdilenTarih"].ToString().Length > 3)
                {
                    label12.Text = dr["TeslimEdilenTarih"].ToString();

                    string tarih2 = Convert.ToDateTime(dr["TeslimTarihi"]).ToString("yyyy-MM-dd HH:mm:ss");

                    DateTime tarihObj = DateTime.ParseExact(tarih, "yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                    DateTime labelTarihObj = DateTime.ParseExact(tarih2, "yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);

                    int gunFarki = (labelTarihObj - tarihObj).Days;

                    label1.Text = dr["TeslimTarihi"].ToString() + "  " + gunFarki.ToString() + " gün önce teslim edildi.";
                }



                if (dr["Aşama"].ToString() == "Hazır")
                {
                    izin = "var";
                    label27.Visible = true;
                    label27.Text = "TESLİM EDİLMEYİ BEKLİYOR..";
                }
                else
                {
                    label27.Visible = true;

                    if (dr["OnayTarihi"].ToString().Length <= 2)
                    {
                        label27.Text = "ONAY BEKLİYOR..";
                    }
                    else if (dr["OnayTarihi"].ToString().Length >= 2)
                    {
                        if (dr["KesildiTarihi"].ToString().Length <= 2)
                        {
                            label27.Text = "CNC'DE BEKLİYOR..";
                        }
                        else if (dr["KesildiTarihi"].ToString().Length >= 2)
                        {
                            label27.Text = "KESİLDİ , " + dr["KesildiTarihi"].ToString();

                            if (dr["Etiket"].ToString().Length <= 2)
                            {
                                label27.Text = "ETİKETTE BEKLİYOR..";
                            }
                            else if (dr["Etiket"].ToString().Length >= 2)
                            {
                                label27.Text = "ETİKET KESİLDİ, PALETTE BEKLİYOR  " + dr["Etiket"].ToString();

                                if (dr["Aşama"].ToString() == "Palet" || dr["Aşama"].ToString() == "Kargo" || dr["Aşama"].ToString() == "Hazır" || dr["Aşama"].ToString() == "Teslim Edildi")
                                {
                                    label27.Text = "ETİKET KESİLDİ, PALETTE BEKLİYOR  " + dr["Etiket"].ToString();

                                    if (dr["MembranPressTarihi"].ToString().Length <= 2)
                                    {
                                        label27.Text = "MEMBRAN PRESSDE BEKLİYOR..";
                                    }
                                    else if (dr["MembranPressTarihi"].ToString().Length >= 2)
                                    {
                                        label27.Text = "PRESS YAPILDI ," + dr["MembranPressTarihi"].ToString();

                                        if (dr["PaketTarihi"].ToString().Length <= 2)
                                        {
                                            label27.Text = "PAKETTE BEKLİYOR..";
                                        }
                                        else if (dr["PaketTarihi"].ToString().Length >= 2)
                                        {
                                            label27.Text = dr["PaketSayısı"] + " PAKET YAPILDI ," + dr["PaketTarihi"].ToString();
                                            if (dr["TeslimEdilenTarih"].ToString().Length <= 2)
                                            {
                                                label27.Text = "TESLİM EDİLMEYİ BEKLİYOR..";
                                            }
                                            else if (dr["TeslimEdilenTarih"].ToString().Length >= 2)
                                            {
                                                label27.Text = "TESLİM EDİLDİ , " + dr["TeslimEdilenTarih"].ToString();
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Lütfen geçerli bir sipariş numarası giriniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                Hesapla();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (izin == "var")
            {
                string sorgu = "UPDATE Siparişler SET Aşama=@Aşama, TeslimEdilenTarih=@TeslimEdilenTarih WHERE SiparisNo=@SiparisNo";
                SqlCommand komut2;
                komut2 = new SqlCommand(sorgu, bgl.baglanti());
                komut2.Parameters.AddWithValue("@SiparisNo", textBox2.Text);
                komut2.Parameters.AddWithValue("@Aşama", "Teslim Edildi");
                komut2.Parameters.AddWithValue("@TeslimEdilenTarih", tarih);
                komut2.ExecuteNonQuery();
                MessageBox.Show(textBox2.Text + " Sipariş teslim edilmiştir.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Hesapla();

                this.Close();
            }
        }

        private void Form19_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (hangiformdan == "form23")
            {
                Form23 yeni = new Form23();
                yeni.yetki = yetki;
                yeni.kullaniciadi = kullaniciadi;
                yeni.hangiformdan = hangiformdan;
                yeni.Show();
                this.Hide();

            }
            else if (hangiformdan == "Form3")
            {
                Form3 frm = new Form3();
                frm.kullaniciadi = kullaniciadi;
                frm.yetki = yetki;
                frm.Show();
            }
            else if(hangiformdan =="form17")
            {
                Form17 frm = new Form17();
                frm.kullaniciadi = kullaniciadi;
                frm.yetki = yetki;
                frm.Show();
            }

        }

        private void Form19_FormClosed(object sender, FormClosedEventArgs e)
        {

        }

        private void Form19_Load(object sender, EventArgs e)
        {

        }
    }
}
