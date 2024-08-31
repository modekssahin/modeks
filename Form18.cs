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
    public partial class Form18 : Form
    {
        public Form18()
        {
            InitializeComponent();
        }
        sqlsinif bgl = new sqlsinif();
        public string yetki;
        public string kullaniciadi;
        public string hangiformdan;
        private void aynısipnugetirme()
        {
            try
            {
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < i; j++)
                    {
                        if (dataGridView1.Rows[j].Cells[1].Value.ToString() == dataGridView1.Rows[i].Cells[1].Value.ToString())
                        {
                            dataGridView1.Rows[i].Visible = false;
                        }
                    }
                }
            }
            catch (Exception)
            {

            }
        }
        private void liste()
        {
            if (textBox1.Text.Length >= 1)
            {
                string srg = textBox1.Text;
                string sorgu = "SELECT Müşteri, SiparisNo From Siparişler where Müşteri Like '" + srg + "%' AND AnaSiparişMi = 'Evet' ORDER BY SiparisNo ASC";
                SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
                DataSet ds = new DataSet();
                adap.Fill(ds, "Siparişler");
                this.dataGridView1.DataSource = ds.Tables[0];
            }
            else if (textBox2.Text.Length >= 1)
            {
                string srg = textBox2.Text;
                string sorgu = "SELECT Müşteri, SiparisNo From Siparişler where SiparisNo Like '" + srg + "' AND AnaSiparişMi = 'Evet' ORDER BY SiparisNo ASC";
                SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
                DataSet ds = new DataSet();
                adap.Fill(ds, "Siparişler");
                this.dataGridView1.DataSource = ds.Tables[0];
            }
            else if (textBox3.Text.Length >= 1)
            {
                string srg = textBox3.Text;
                string sorgu = "SELECT Müşteri, SiparisNo From Siparişler where BID Like '" + srg + "' AND AnaSiparişMi = 'Evet' ORDER BY SiparisNo ASC";
                SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
                DataSet ds = new DataSet();
                adap.Fill(ds, "Siparişler");
                this.dataGridView1.DataSource = ds.Tables[0];
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            liste();
            aynısipnugetirme();
        }
        private void SipNoyaGöreBilgileriGetir()
        {
            label1.BackColor = Color.Green;
            liste();
            aynısipnugetirme();
            aşama = "";
            paketsayısı = "";
            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT * FROM Siparişler where SiparisNo=@p1";
            komut.Parameters.AddWithValue("@p1", textBox2.Text);
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                label46.Text = dr["SiparişTarihi"].ToString();
                label1.Text = dr["OnayTarihi"].ToString();
                label5.Text = dr["TeslimTarihi"].ToString();
                label7.Text = dr["KesildiTarihi"].ToString();
                label9.Text = dr["Etiket"].ToString();
                label11.Text = dr["MembranPressTarihi"].ToString();
                label13.Text = dr["PaketTarihi"].ToString();
                label15.Text = dr["TeslimEdilenTarih"].ToString();
                aşama = dr["Aşama"].ToString();
                paketsayısı = dr["PaketSayısı"].ToString();

                if (dr["OnayTarihi"].ToString().Length <= 2)
                {
                    label7.Text = "ONAY BEKLİYOR..";
                    label9.Text = "ONAY BEKLİYOR..";
                    label11.Text = "ONAY BEKLİYOR..";
                    label13.Text = "ONAY BEKLİYOR..";
                    label15.Text = "ONAY BEKLİYOR..";
                    label15.Text = "ONAY BEKLİYOR..";
                }
                else if (dr["OnayTarihi"].ToString().Length >= 2)
                {
                    if (dr["KesildiTarihi"].ToString().Length <= 2)
                    {
                        label7.Text = "CNC'DE BEKLİYOR..";
                        label9.Text = "CNC'DE BEKLİYOR..";
                        label11.Text = "CNC'DE BEKLİYOR..";
                        label13.Text = "CNC'DE BEKLİYOR..";
                        label15.Text = "CNC'DE BEKLİYOR..";
                    }
                    else if (dr["KesildiTarihi"].ToString().Length >= 2)
                    {
                        label7.Text = "KESİLDİ , " + dr["KesildiTarihi"].ToString();
                        label7.BackColor = Color.Green;

                        if (dr["Etiket"].ToString().Length <= 2)
                        {
                            label9.Text = "ETİKETTE BEKLİYOR..";
                            label11.Text = "ETİKETTE BEKLİYOR..";
                            label13.Text = "ETİKETTE BEKLİYOR..";
                            label15.Text = "ETİKETTE BEKLİYOR..";
                        }
                        else if (dr["Etiket"].ToString().Length >= 2)
                        {
                            label9.Text = "ETİKET KESİLDİ ," + dr["Etiket"].ToString();
                            label9.BackColor = Color.Green;
                            label11.Text = "PALETTE BEKLİYOR..";
                            label13.Text = "PALETTE BEKLİYOR..";
                            label15.Text = "PALETTE BEKLİYOR..";

                            if (aşama == "Palet" || aşama == "Kargo" || aşama == "Hazır" || aşama == "Teslim Edildi")
                            {
                                label9.Text = "ETİKET KESİLDİ ," + dr["Etiket"].ToString();
                                label9.BackColor = Color.Green;

                                if (dr["MembranPressTarihi"].ToString().Length <= 2)
                                {
                                    label11.Text = "MEMBRAN PRESSDE BEKLİYOR..";
                                    label13.Text = "MEMBRAN PRESSDE BEKLİYOR..";
                                    label15.Text = "MEMBRAN PRESSDE BEKLİYOR..";
                                }
                                else if (dr["MembranPressTarihi"].ToString().Length >= 2)
                                {
                                    label11.Text = "PRESS YAPILDI ," + dr["MembranPressTarihi"].ToString();
                                    label11.BackColor = Color.Green;

                                    if (dr["PaketTarihi"].ToString().Length <= 2)
                                    {
                                        label13.Text = "PAKETTE BEKLİYOR..";
                                        label15.Text = "PAKETTE BEKLİYOR..";
                                    }
                                    else if (dr["PaketTarihi"].ToString().Length >= 2)
                                    {
                                        label13.Text = paketsayısı + " PAKET YAPILDI ," + dr["PaketTarihi"].ToString();
                                        label13.BackColor = Color.Green;
                                        if (dr["TeslimEdilenTarih"].ToString().Length <= 2)
                                        {
                                            label15.Text = "TESLİM EDİLMEYİ BEKLİYOR..";
                                        }
                                        else if (dr["TeslimEdilenTarih"].ToString().Length >= 2)
                                        {
                                            label15.Text = "TESLİM EDİLDİ , " + dr["TeslimEdilenTarih"].ToString();
                                            label15.BackColor = Color.Green;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

            }
        }
        string aşama;
        string paketsayısı;
        private void button2_Click(object sender, EventArgs e)
        {
            SipNoyaGöreBilgileriGetir();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            aynısipnugetirme();
        }

        private void Form18_Load(object sender, EventArgs e)
        {

        }
        private void BIDeGöreBilgileriGetir()
        {
            label1.BackColor = Color.Green;
            liste();
            aynısipnugetirme();
            aşama = "";
            paketsayısı = "";
            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT * FROM Siparişler where BID=@p1";
            komut.Parameters.AddWithValue("@p1", textBox3.Text);
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                label46.Text = dr["SiparişTarihi"].ToString();
                label1.Text = dr["OnayTarihi"].ToString();
                label5.Text = dr["TeslimTarihi"].ToString();
                label7.Text = dr["KesildiTarihi"].ToString();
                label9.Text = dr["Etiket"].ToString();
                label11.Text = dr["MembranPressTarihi"].ToString();
                label13.Text = dr["PaketTarihi"].ToString();
                label15.Text = dr["TeslimEdilenTarih"].ToString();
                aşama = dr["Aşama"].ToString();
                paketsayısı = dr["PaketSayısı"].ToString();

                if (dr["OnayTarihi"].ToString().Length <= 2)
                {
                    label7.Text = "ONAY BEKLİYOR..";
                    label9.Text = "ONAY BEKLİYOR..";
                    label11.Text = "ONAY BEKLİYOR..";
                    label13.Text = "ONAY BEKLİYOR..";
                    label15.Text = "ONAY BEKLİYOR..";
                    label15.Text = "ONAY BEKLİYOR..";
                }
                else if (dr["OnayTarihi"].ToString().Length >= 2)
                {
                    if (dr["KesildiTarihi"].ToString().Length <= 2)
                    {
                        label7.Text = "CNC'DE BEKLİYOR..";
                        label9.Text = "CNC'DE BEKLİYOR..";
                        label11.Text = "CNC'DE BEKLİYOR..";
                        label13.Text = "CNC'DE BEKLİYOR..";
                        label15.Text = "CNC'DE BEKLİYOR..";
                    }
                    else if (dr["KesildiTarihi"].ToString().Length >= 2)
                    {
                        label7.Text = "KESİLDİ , " + dr["KesildiTarihi"].ToString();
                        label7.BackColor = Color.Green;

                        if (dr["Etiket"].ToString().Length <= 2)
                        {
                            label9.Text = "ETİKETTE BEKLİYOR..";
                            label11.Text = "ETİKETTE BEKLİYOR..";
                            label13.Text = "ETİKETTE BEKLİYOR..";
                            label15.Text = "ETİKETTE BEKLİYOR..";
                        }
                        else if (dr["Etiket"].ToString().Length >= 2)
                        {
                            label9.Text = "ETİKET KESİLDİ ," + dr["Etiket"].ToString();
                            label9.BackColor = Color.Green;
                            label11.Text = "PALETTE BEKLİYOR..";
                            label13.Text = "PALETTE BEKLİYOR..";
                            label15.Text = "PALETTE BEKLİYOR..";

                            if (aşama == "Palet" || aşama == "Kargo" || aşama == "Hazır" || aşama == "Teslim Edildi")
                            {
                                label9.Text = "ETİKET KESİLDİ ," + dr["Etiket"].ToString();
                                label9.BackColor = Color.Green;

                                if (dr["MembranPressTarihi"].ToString().Length <= 2)
                                {
                                    label11.Text = "MEMBRAN PRESSDE BEKLİYOR..";
                                    label13.Text = "MEMBRAN PRESSDE BEKLİYOR..";
                                    label15.Text = "MEMBRAN PRESSDE BEKLİYOR..";
                                }
                                else if (dr["MembranPressTarihi"].ToString().Length >= 2)
                                {
                                    label11.Text = "PRESS YAPILDI ," + dr["MembranPressTarihi"].ToString();
                                    label11.BackColor = Color.Green;

                                    if (dr["PaketTarihi"].ToString().Length <= 2)
                                    {
                                        label13.Text = "PAKETTE BEKLİYOR..";
                                        label15.Text = "PAKETTE BEKLİYOR..";
                                    }
                                    else if (dr["PaketTarihi"].ToString().Length >= 2)
                                    {
                                        label13.Text = paketsayısı + " PAKET YAPILDI ," + dr["PaketTarihi"].ToString();
                                        label13.BackColor = Color.Green;
                                        if (dr["TeslimEdilenTarih"].ToString().Length <= 2)
                                        {
                                            label15.Text = "TESLİM EDİLMEYİ BEKLİYOR..";
                                        }
                                        else if (dr["TeslimEdilenTarih"].ToString().Length >= 2)
                                        {
                                            label15.Text = "TESLİM EDİLDİ , " + dr["TeslimEdilenTarih"].ToString();
                                            label15.BackColor = Color.Green;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            BIDeGöreBilgileriGetir();
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                liste();
                aynısipnugetirme();
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                SipNoyaGöreBilgileriGetir();
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                BIDeGöreBilgileriGetir();
            }
        }

        private void Form18_FormClosed(object sender, FormClosedEventArgs e)
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
            else if (hangiformdan == "form17")
            {
                Form17 frm = new Form17();
                frm.kullaniciadi = kullaniciadi;
                frm.yetki = yetki;
                frm.Show();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text == "" && textBox2.Text == "" && textBox3.Text == "")
            {
                string srg = textBox2.Text;
                string sorgu = "SELECT Müşteri, SiparisNo From Siparişler where SiparisNo Like '" + srg + "' AND AnaSiparişMi = 'Evet' ORDER BY SiparisNo ASC";
                SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
                DataSet ds = new DataSet();
                adap.Fill(ds, "Siparişler");
                this.dataGridView1.DataSource = ds.Tables[0];
            }
        }
    }
}
