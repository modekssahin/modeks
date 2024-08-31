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
using FirebirdSql.Data.FirebirdClient;
using Microsoft.VisualBasic;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Modeks
{
    public partial class Form8 : Form
    {
        public Form8()
        {
            InitializeComponent();
        }
        sqlsinif bgl = new sqlsinif();
        public string yetki;
        double basılacakmatm2 = 0;
        double basılacakmatadet = 0;
        double basılacakmatacilm2 = 0;
        double basılacakmataciladet = 0;

        double basılacakparlakm2 = 0;
        double basılacakparlakadet = 0;
        double basılacakparlakacilm2 = 0;
        double basılacakparlakaciladet = 0;

        double basılacaktoplamm2 = 0;
        double basılacaktoplamadet = 0;

        double bugünbasılanmatm2 = 0;
        double bugünbasılanmatadet = 0;
        double bugünbasılanmatacilm2 = 0;
        double bugünbasılanmataciladet = 0;

        double bugünbasılanparlakm2 = 0;
        double bugünbasılanparlakadet = 0;
        double bugünbasılanparlakacilm2 = 0;
        double bugünbasılanparlakaciladet = 0;

        double bugünbasılantoplamm2 = 0;
        double bugünbasılantoplamadet = 0;

        private void liste()
        {
            string kayit = "SELECT DISTINCT SiparisNo,Müşteri,SiparişTipi,Renk,ToplamM2,ToplamAdet From Siparişler Where Aşama=@Aşama AND Renk LIKE 'HG%' AND MembranPressTarihi IS NULL ORDER BY SiparisNo DESC";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            komut.Parameters.AddWithValue("@Aşama", "Palet");
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView2.DataSource = dt;
        }

        private void liste_üretim_formu()
        {
            string kayit = "SELECT DISTINCT SiparisNo,Müşteri,SiparişTipi,Renk,ToplamM2,ToplamAdet,SiparişTarihi,TeslimTarihi,Adres,Telefon From Siparişler Where Aşama=@Aşama AND Renk LIKE 'HG%' AND MembranPressTarihi IS NULL";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            komut.Parameters.AddWithValue("@Aşama", "Palet");
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView6.DataSource = dt;
        }
        private void liste2()
        {
            string kayit = "SELECT DISTINCT SiparisNo,Müşteri,SiparişTipi,Renk,ToplamM2,ToplamAdet From Siparişler Where Aşama=@Aşama AND Renk NOT LIKE 'HG%' AND MembranPressTarihi IS NULL ORDER BY SiparisNo DESC";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            komut.Parameters.AddWithValue("@Aşama", "Palet");
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
        }
        private void liste2_üretim_formu()
        {
            string kayit = "SELECT DISTINCT SiparisNo,Müşteri,SiparişTipi,Renk,ToplamM2,ToplamAdet,SiparişTarihi,TeslimTarihi,Adres,Telefon From Siparişler Where Aşama=@Aşama AND Renk NOT LIKE 'HG%' AND MembranPressTarihi IS NULL";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            komut.Parameters.AddWithValue("@Aşama", "Palet");
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView5.DataSource = dt;
        }
        private void BasılacakMatM2()
        {
            basılacakmatm2 = 0;
            basılacakmatacilm2 = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true)
                    if (dataGridView1.Rows[i].Cells["SiparişTipi"].Value.ToString() == "Normal" || dataGridView1.Rows[i].Cells["SiparişTipi"].Value.ToString() == "Üretim Sorunu")
                    {
                        basılacakmatm2 += Convert.ToDouble(dataGridView1.Rows[i].Cells["ToplamM2"].Value);
                    }
                    else if (dataGridView1.Rows[i].Cells["SiparişTipi"].Value.ToString() == "Acil")
                    {
                        basılacakmatacilm2 += Convert.ToDouble(dataGridView1.Rows[i].Cells["ToplamM2"].Value);
                    }

            }
            textBox1.Text = basılacakmatm2.ToString("0.##");
            textBox6.Text = basılacakmatacilm2.ToString("0.##");
        }
        private void BasılacakMatAdet()
        {
            basılacakmatadet = 0;
            basılacakmataciladet = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true)
                    if (dataGridView1.Rows[i].Cells["SiparişTipi"].Value.ToString() == "Normal" || dataGridView1.Rows[i].Cells["SiparişTipi"].Value.ToString() == "Üretim Sorunu")
                    {
                        basılacakmatadet += Convert.ToDouble(dataGridView1.Rows[i].Cells["ToplamAdet"].Value);
                    }
                    else if (dataGridView1.Rows[i].Cells["SiparişTipi"].Value.ToString() == "Acil")
                    {
                        basılacakmataciladet += Convert.ToDouble(dataGridView1.Rows[i].Cells["ToplamAdet"].Value);
                    }
            }
            textBox2.Text = basılacakmatadet.ToString("0.##");
            textBox5.Text = basılacakmataciladet.ToString("0.##");
        }
        private void BasılacakParlakM2()
        {
            basılacakparlakm2 = 0;
            basılacakparlakacilm2 = 0;
            for (int i = 0; i < dataGridView2.Rows.Count - 1; ++i)
            {
                if (dataGridView2.Rows[i].Visible == true)
                    if (dataGridView1.Rows[i].Cells["SiparişTipi"].Value.ToString() == "Normal" || dataGridView1.Rows[i].Cells["SiparişTipi"].Value.ToString() == "Üretim Sorunu")
                    {
                        basılacakparlakm2 += Convert.ToDouble(dataGridView2.Rows[i].Cells["ToplamM2"].Value);
                    }
                    else if (dataGridView2.Rows[i].Cells["SiparişTipi"].Value.ToString() == "Acil")
                    {
                        basılacakparlakacilm2 += Convert.ToDouble(dataGridView2.Rows[i].Cells["ToplamM2"].Value);
                    }
            }
            textBox16.Text = basılacakparlakm2.ToString("0.##");
            textBox4.Text = basılacakparlakacilm2.ToString("0.##");
        }
        private void BasılacakParlakAdet()
        {
            basılacakparlakadet = 0;
            basılacakparlakaciladet = 0;
            for (int i = 0; i < dataGridView2.Rows.Count - 1; ++i)
            {
                if (dataGridView2.Rows[i].Visible == true)
                    if (dataGridView1.Rows[i].Cells["SiparişTipi"].Value.ToString() == "Normal" || dataGridView1.Rows[i].Cells["SiparişTipi"].Value.ToString() == "Üretim Sorunu")
                    {
                        basılacakparlakadet += Convert.ToDouble(dataGridView2.Rows[i].Cells["ToplamAdet"].Value);
                    }
                    else if (dataGridView2.Rows[i].Cells["SiparişTipi"].Value.ToString() == "Acil")
                    {
                        basılacakparlakaciladet += Convert.ToDouble(dataGridView2.Rows[i].Cells["ToplamAdet"].Value);
                    }
            }
            textBox15.Text = basılacakparlakadet.ToString("0.##");
            textBox3.Text = basılacakparlakaciladet.ToString("0.##");
        }
        private void BasılacakToplamM2()
        {
            basılacaktoplamm2 = 0;
            basılacaktoplamm2 = basılacakmatm2 + basılacakparlakm2 + basılacakmatacilm2 + basılacakparlakacilm2;
            textBox18.Text = basılacaktoplamm2.ToString("0.##");
        }
        private void BasılacakToplamAdet()
        {
            basılacaktoplamadet = 0;
            basılacaktoplamadet = basılacakmatadet + basılacakparlakadet + basılacakmataciladet + basılacakparlakaciladet;
            textBox17.Text = basılacaktoplamadet.ToString("0.##");
        }

        int kayıtsayısı;
        string renk;
        string sipariştipi;
        DateTime tarih;
        private void BugünBasılanlar()
        {
            bugünbasılanmatm2 = 0;
            bugünbasılanmatadet = 0;

            bugünbasılanmatacilm2 = 0;
            bugünbasılanmataciladet = 0;

            bugünbasılanparlakm2 = 0;
            bugünbasılanparlakadet = 0;

            bugünbasılanparlakacilm2 = 0;
            bugünbasılanparlakaciladet = 0;

            bugünbasılantoplamm2 = 0;
            bugünbasılantoplamadet = 0;
            SqlCommand cmd = new SqlCommand("select count(*) from Siparişler Where Aşama=@p1", bgl.baglanti());
            cmd.Parameters.AddWithValue("@p1", "Kargo");
            kayıtsayısı = Convert.ToInt32(cmd.ExecuteScalar());

            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT *FROM Siparişler where AnaSiparişMi=@AnaSiparişMi";
            komut.Parameters.AddWithValue("@AnaSiparişMi", "Evet");
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                renk = dr["Renk"].ToString();
                sipariştipi = dr["SiparişTipi"].ToString();
                if (dr["MembranPressTarihi"].ToString() != "")
                {
                    tarih = Convert.ToDateTime(dr["MembranPressTarihi"].ToString());
                    if (renk.StartsWith("HG") && tarih.ToString("yyyy - MM - dd") == (DateTime.Now.ToString("yyyy - MM - dd")))
                    {
                        if (sipariştipi == "Normal" || sipariştipi == "Üretim Sorunu")
                        {
                            bugünbasılanparlakm2 += Convert.ToDouble(dr["M2"]);
                            bugünbasılanparlakadet += Convert.ToDouble(dr["Adet"]);

                            textBox22.Text = bugünbasılanparlakm2.ToString();
                            textBox21.Text = bugünbasılanparlakadet.ToString();
                        }
                        else if (sipariştipi == "Acil")
                        {
                            bugünbasılanparlakacilm2 += Convert.ToDouble(dr["M2"]);
                            bugünbasılanparlakaciladet += Convert.ToDouble(dr["Adet"]);

                            textBox8.Text = bugünbasılanparlakacilm2.ToString();
                            textBox7.Text = bugünbasılanparlakaciladet.ToString();
                        }


                    }
                    else if (!renk.StartsWith("HG") && tarih.ToString("yyyy - MM - dd") == (DateTime.Now.ToString("yyyy - MM - dd")))
                    {
                        if (sipariştipi == "Normal" || sipariştipi == "Üretim Sorunu")
                        {
                            bugünbasılanmatm2 += Convert.ToDouble(dr["M2"]);
                            bugünbasılanmatadet += Convert.ToDouble(dr["Adet"]);
                            textBox24.Text = bugünbasılanmatm2.ToString();
                            textBox23.Text = bugünbasılanmatadet.ToString();
                        }
                        else if (sipariştipi == "Acil")
                        {
                            bugünbasılanmatacilm2 += Convert.ToDouble(dr["M2"]);
                            bugünbasılanmataciladet += Convert.ToDouble(dr["Adet"]);

                            textBox10.Text = bugünbasılanmatacilm2.ToString();
                            textBox9.Text = bugünbasılanmataciladet.ToString();
                        }
                    }
                }
            }
            bugünbasılantoplamm2 += bugünbasılanmatm2 + bugünbasılanparlakm2 + bugünbasılanmatacilm2 + bugünbasılanparlakacilm2;
            bugünbasılantoplamadet += bugünbasılanmatadet + bugünbasılanparlakadet + bugünbasılanmataciladet + bugünbasılanparlakaciladet;
            textBox33.Text = bugünbasılantoplamm2.ToString();
            textBox20.Text = bugünbasılantoplamm2.ToString();
            textBox34.Text = bugünbasılantoplamadet.ToString();
            textBox19.Text = bugünbasılantoplamadet.ToString();
            textBox35.Text = (bugünbasılantoplamadet / bugünbasılantoplamm2).ToString();
        }

        public void groupBox1_Paint(object sender, PaintEventArgs e)
        {
            GroupBox box = sender as GroupBox;
            DrawGroupBox(box, e.Graphics, Color.Black, Color.Black);
        }
        public void DrawGroupBox(GroupBox box, Graphics g, Color textColor, Color borderColor)
        {
            if (box != null)
            {
                Brush textBrush = new SolidBrush(textColor);
                Brush borderBrush = new SolidBrush(borderColor);
                Pen borderPen = new Pen(borderBrush);
                SizeF strSize = g.MeasureString(box.Text, box.Font);
                Rectangle rect = new Rectangle(box.ClientRectangle.X,
                                               box.ClientRectangle.Y + (int)(strSize.Height / 2),
                                               box.ClientRectangle.Width - 1,
                                               box.ClientRectangle.Height - (int)(strSize.Height / 2) - 1);
                // Clear text and border
                g.Clear(this.BackColor);
                // Draw text
                g.DrawString(box.Text, box.Font, textBrush, box.Padding.Left, 0);
                // Drawing Border
                //Left
                g.DrawLine(borderPen, rect.Location, new Point(rect.X, rect.Y + rect.Height));
                //Right
                g.DrawLine(borderPen, new Point(rect.X + rect.Width, rect.Y), new Point(rect.X + rect.Width, rect.Y + rect.Height));
                //Bottom
                g.DrawLine(borderPen, new Point(rect.X, rect.Y + rect.Height), new Point(rect.X + rect.Width, rect.Y + rect.Height));
                //Top1
                g.DrawLine(borderPen, new Point(rect.X, rect.Y), new Point(rect.X + box.Padding.Left, rect.Y));
                //Top2
                g.DrawLine(borderPen, new Point(rect.X + box.Padding.Left + (int)(strSize.Width), rect.Y), new Point(rect.X + rect.Width, rect.Y));
            }
        }

        private void Methodlar()
        {
            liste();
            liste2();
            liste_üretim_formu();
            liste2_üretim_formu();
            aynısipnugetirme();
            aynısipnugetirme2();

            BasılacakMatM2();
            BasılacakMatAdet();
            BasılacakParlakM2();
            BasılacakParlakAdet();
            BasılacakToplamM2();
            BasılacakToplamAdet();
            BugünBasılanlar();
            AcilSipariş();
            TeslimTarihine3GünKalanlarıYakSöndür();
            siparisnogetir();
            renkgetir();


        }

        double toplam2 = 0;
        private void AşağıdakiDatagridMat()
        {
            if (dataGridView3.Columns.Count == 2)
                dataGridView3.Columns.RemoveAt(1);
            string kayit = "SELECT DISTINCT Renk From Siparişler Where Aşama=@Aşama AND Renk NOT LIKE 'HG%' AND MembranPressTarihi IS NULL";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            komut.Parameters.AddWithValue("@Aşama", "Palet");
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView3.DataSource = dt;
            dataGridView3.Columns.Add("Column", "Toplam");
            for (int i = 0; i < dataGridView3.Rows.Count - 1; i++)
            {
                toplam2 = 0;
                for (int y = 0; y < dataGridView1.Rows.Count - 1; y++)
                {
                    if (dataGridView3.Rows[i].Cells[0].Value.ToString() == dataGridView1.Rows[y].Cells["Renk"].Value.ToString())
                    {
                        toplam2 += Convert.ToDouble(dataGridView1.Rows[y].Cells["ToplamM2"].Value);
                        dataGridView3.Rows[i].Cells[1].Value = toplam2.ToString();
                    }
                }
            }
        }
        double toplam = 0;
        private void AşağıdakiDatagridParlak()
        {
            if (dataGridView4.Columns.Count == 2)
                dataGridView4.Columns.RemoveAt(1);
            string kayit = "SELECT DISTINCT Renk From Siparişler Where Aşama=@Aşama AND Renk LIKE 'HG%' AND MembranPressTarihi IS NULL";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            komut.Parameters.AddWithValue("@Aşama", "Palet");
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView4.DataSource = dt;

            dataGridView4.Columns.Add("Column", "Toplam");

            for (int i = 0; i < dataGridView4.Rows.Count - 1; i++)
            {
                toplam = 0;
                for (int y = 0; y < dataGridView2.Rows.Count - 1; y++)
                {
                    if (dataGridView4.Rows[i].Cells[0].Value.ToString() == dataGridView2.Rows[y].Cells["Renk"].Value.ToString())
                    {
                        toplam += Convert.ToDouble(dataGridView2.Rows[y].Cells["ToplamM2"].Value);
                        dataGridView4.Rows[i].Cells[1].Value = toplam.ToString();
                    }

                }
            }
        }

        string acilsipariş;
        string kesilditarihi;
        string Etiket;
        string membrantarihi;
        private void AcilSipariş()
        {
            kesilditarihi = "";
            acilsipariş = "";

            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT *FROM Siparişler where Onay=@Onay AND AnaSiparişMi=@p1";
            komut.Parameters.AddWithValue("@Onay", "Onaylandı");
            komut.Parameters.AddWithValue("@p1", "Evet");
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                acilsipariş = dr["SiparişTipi"].ToString();
                kesilditarihi = dr["KesildiTarihi"].ToString();
                Etiket = dr["Etiket"].ToString();
                membrantarihi = dr["MembranPressTarihi"].ToString();
                if ((acilsipariş == "Acil" && kesilditarihi.Length < 2) || (acilsipariş == "Acil" && Etiket.Length < 2) || (acilsipariş == "Acil" && membrantarihi.Length < 2))
                {
                    timer2.Start();
                    label9.Text = "Dikkat! Acil Sipariş Var! Dikkat! Acil Sipariş Var! Dikkat! Acil Sipariş Var!";
                    label9.BackColor = Color.Red;
                    break;
                }
                else
                {
                    timer2.Stop();
                    label9.Text = "İSTATİSTİKLER";
                    label9.TextAlign = ContentAlignment.MiddleCenter;
                    label9.BackColor = Color.Blue;
                }
            }
        }
        private void siparisnogetir()
        {
            comboBox1.Items.Clear();
            comboBox3.Items.Clear();
            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT DISTINCT SiparisNo FROM Siparişler where Aşama='" + "Palet" + "'  AND MembranPressTarihi IS NULL ORDER BY SiparisNo ASC";
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                comboBox1.Items.Add(dr["SiparisNo"]);
                comboBox3.Items.Add(dr["SiparisNo"]);
            }
        }
        private void aynısipnugetirme()
        {
            try
            {
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < i; j++)
                    {
                        if (dataGridView1.Rows[j].Cells[0].Value.ToString() == dataGridView1.Rows[i].Cells[0].Value.ToString())
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
        private void aynısipnugetirme2()
        {
            try
            {
                for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < i; j++)
                    {
                        if (dataGridView2.Rows[j].Cells[0].Value.ToString() == dataGridView2.Rows[i].Cells[0].Value.ToString())
                        {
                            dataGridView2.Rows[i].Visible = false;
                        }

                    }

                }
            }
            catch (Exception)
            {

            }
        }
        private void Form8_Load(object sender, EventArgs e)
        {
            if (yetki == "Membran / Press")
            {
                button28.Visible = false;
            }
            timer1.Start();
            Methodlar();
            AşağıdakiDatagridMat();
            AşağıdakiDatagridParlak();
        }
        private void siparisnoyagöresıralahg()
        {
            string srg = comboBox1.Text;
            string sorgu = "SELECT SiparisNo,SiparişTipi,Renk,ToplamM2,ToplamAdet From Siparişler where SiparisNo Like '" + srg + "' AND Renk LIKE 'HG%' AND Aşama='" + "Palet" + "' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView2.DataSource = ds.Tables[0];
        }
        private void siparisnoyagöresıralanothg()
        {
            string srg = comboBox1.Text;
            string sorgu = "SELECT SiparisNo,SiparişTipi,Renk,ToplamM2,ToplamAdet From Siparişler where SiparisNo Like '" + srg + "' AND Renk NOT LIKE 'HG%' AND Aşama='" + "Palet" + "' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
        }
        private void renkgetir()
        {
            comboBox2.Items.Clear();
            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT DISTINCT Renk FROM Siparişler where Aşama='" + "Palet" + "' AND MembranPressTarihi IS NULL ORDER BY Renk ASC";
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                comboBox2.Items.Add(dr["Renk"]);
            }
        }
        private void rengegöresıralahg()
        {
            string srg = comboBox2.Text;
            string sorgu = "SELECT DISTINCT SiparisNo,SiparişTipi,Renk,ToplamM2,ToplamAdet From Siparişler where Renk Like '" + srg + "' AND Renk LIKE 'HG%' AND Aşama='" + "Palet" + "' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView2.DataSource = ds.Tables[0];
        }
        private void rengegöresıralanothg()
        {
            string srg = comboBox2.Text;
            string sorgu = "SELECT DISTINCT SiparisNo,SiparişTipi,Renk,ToplamM2,ToplamAdet From Siparişler where Renk Like '" + srg + "' AND Renk NOT LIKE 'HG%' AND Aşama='" + "Palet" + "' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
        }
        private void pictureBox8_Click(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            label2.Text = DateTime.Now.ToLongDateString();
            label12.Text = DateTime.Now.ToLongTimeString();
        }

        private void PaketeGönder()
        {
            if (comboBox3.Text != "")
            {
                if (dataGridView2.Rows.Count > dataGridView1.Rows.Count)
                {
                    for (int a = 0; a < dataGridView2.Rows.Count - 1; a++)
                    {
                        if (dataGridView2.Rows.Count - 1 > a && comboBox3.Text == dataGridView2.Rows[a].Cells[0].Value.ToString())
                        {
                            string tarih = Convert.ToDateTime(DateTime.Now).ToString("yyyy-MM-dd HH:mm:ss");

                            string sorgu = "UPDATE Siparişler SET Aşama=@Aşama, MembranPressTarihi=@MembranPressTarihi WHERE SiparisNo=@SiparisNo";
                            SqlCommand komut;
                            komut = new SqlCommand(sorgu, bgl.baglanti());
                            komut.Parameters.AddWithValue("@SiparisNo", comboBox3.Text);
                            komut.Parameters.AddWithValue("@Aşama", "Kargo");
                            komut.Parameters.AddWithValue("@MembranPressTarihi", tarih);
                            komut.ExecuteNonQuery();
                            MessageBox.Show("Sipariş Kargoya Gönderilmiştir", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            label19.Text = comboBox3.Text;
                            label25.Text = dataGridView2.Rows[0].Cells["Müşteri"].Value.ToString();

                            Methodlar();
                            AşağıdakiDatagridParlak();
                            break;

                        }
                        else if (dataGridView1.Rows.Count - 1 > a && comboBox3.Text == dataGridView1.Rows[a].Cells[0].Value.ToString())
                        {

                            string tarih = Convert.ToDateTime(DateTime.Now).ToString("yyyy-MM-dd HH:mm:ss");

                            string sorgu = "UPDATE Siparişler SET Aşama=@Aşama, MembranPressTarihi=@MembranPressTarihi WHERE SiparisNo=@SiparisNo";
                            SqlCommand komut;
                            komut = new SqlCommand(sorgu, bgl.baglanti());
                            komut.Parameters.AddWithValue("@SiparisNo", comboBox3.Text);
                            komut.Parameters.AddWithValue("@Aşama", "Kargo");
                            komut.Parameters.AddWithValue("@MembranPressTarihi", tarih);
                            komut.ExecuteNonQuery();
                            MessageBox.Show("Sipariş Kargoya Gönderilmiştir", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            label19.Text = comboBox3.Text;
                            label25.Text = dataGridView1.Rows[0].Cells["Müşteri"].Value.ToString();
                            Methodlar();
                            AşağıdakiDatagridMat();
                            break;

                        }
                        else
                        {
                            if (a == dataGridView2.Rows.Count - 1)
                            {
                                MessageBox.Show("Geçerli bir sipariş numarası girin!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                break;
                            }

                        }

                    }
                }
                else
                {
                    for (int b = 0; b < dataGridView1.Rows.Count - 1; b++)
                    {
                        if (dataGridView1.Rows.Count - 1 > b && comboBox3.Text == dataGridView1.Rows[b].Cells[0].Value.ToString())
                        {
                            string tarih = Convert.ToDateTime(DateTime.Now).ToString("yyyy-MM-dd HH:mm:ss");

                            string sorgu = "UPDATE Siparişler SET Aşama=@Aşama, MembranPressTarihi=@MembranPressTarihi WHERE SiparisNo=@SiparisNo";
                            SqlCommand komut;
                            komut = new SqlCommand(sorgu, bgl.baglanti());
                            komut.Parameters.AddWithValue("@SiparisNo", comboBox3.Text);
                            komut.Parameters.AddWithValue("@Aşama", "Kargo");
                            komut.Parameters.AddWithValue("@MembranPressTarihi", tarih);
                            komut.ExecuteNonQuery();
                            MessageBox.Show("Sipariş Kargoya Gönderilmiştir", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            label19.Text = comboBox3.Text;
                            label25.Text = dataGridView1.Rows[0].Cells["Müşteri"].Value.ToString();
                            Methodlar();
                            AşağıdakiDatagridMat();
                            break;
                        }
                        else if (dataGridView2.Rows.Count - 1 > b && comboBox3.Text == dataGridView2.Rows[b].Cells[0].Value.ToString())
                        {
                            string tarih = Convert.ToDateTime(DateTime.Now).ToString("yyyy-MM-dd HH:mm:ss");

                            string sorgu = "UPDATE Siparişler SET Aşama=@Aşama, MembranPressTarihi=@MembranPressTarihi WHERE SiparisNo=@SiparisNo";
                            SqlCommand komut;
                            komut = new SqlCommand(sorgu, bgl.baglanti());
                            komut.Parameters.AddWithValue("@SiparisNo", comboBox3.Text);
                            komut.Parameters.AddWithValue("@Aşama", "Kargo");
                            komut.Parameters.AddWithValue("@MembranPressTarihi", tarih);
                            komut.ExecuteNonQuery();
                            MessageBox.Show("Sipariş Kargoya Gönderilmiştir", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            label19.Text = comboBox3.Text;
                            label25.Text = dataGridView2.Rows[0].Cells["Müşteri"].Value.ToString();
                            Methodlar();
                            AşağıdakiDatagridParlak();
                            break;
                        }
                        else
                        {
                            if (b == dataGridView1.Rows.Count - 1)
                            {
                                MessageBox.Show("Geçerli bir sipariş numarası girin!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                break;
                            }
                        }

                    }
                }

            }
            else
            {
                MessageBox.Show("Bir sipariş numarası girin!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            PaketeGönder();
            Methodlar();
            comboBox3.Text = "";
        }

        private void comboBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                PaketeGönder();
                Methodlar();
                comboBox3.Text = "";
            }
        }
        string siparisnoacil, renkacil;
        private void TeslimTarihine3GünKalanlarıYakSöndür()
        {
            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT * FROM Siparişler WHERE CONVERT(datetime, TeslimTarihi, 104) <= DATEADD(day, 6, GETDATE()) AND Onay=@Onay AND AnaSiparişMi=@p1 AND MembranPressTarihi is null ORDER BY SiparisNo ASC";
            komut.Parameters.AddWithValue("@Onay", "Onaylandı");
            komut.Parameters.AddWithValue("@p1", "Evet");
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                siparisnoacil = dr["SiparisNo"].ToString();
                renkacil = dr["Renk"].ToString();
                label9.Text += " / " + siparisnoacil + "-" + renkacil + " / ";
            }
        }
        private void timer2_Tick(object sender, EventArgs e)
        {
            label9.Text = label9.Text.Substring(1) + label9.Text.Substring(0, 1);
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            siparisnoyagöresıralahg();
            siparisnoyagöresıralanothg();
            aynısipnugetirme();
            aynısipnugetirme2();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            rengegöresıralahg();
            rengegöresıralanothg();
            aynısipnugetirme();
            aynısipnugetirme2();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            liste();
            liste2();
            aynısipnugetirme();
            aynısipnugetirme2();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            string sorgu = "SELECT DISTINCT SiparisNo,Müşteri,SiparişTipi,Renk,ToplamM2,ToplamAdet From Siparişler where Renk LIKE 'HG%' AND Aşama='" + "Palet" + "' ORDER BY Renk DESC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView2.DataSource = ds.Tables[0];

            string sorgu2 = "SELECT DISTINCT SiparisNo,Müşteri,SiparişTipi,Renk,ToplamM2,ToplamAdet From Siparişler where Renk NOT LIKE 'HG%' AND Aşama='" + "Palet" + "' ORDER BY Renk DESC";
            SqlDataAdapter adap2 = new SqlDataAdapter(sorgu2, bgl.baglanti());
            DataSet ds2 = new DataSet();
            adap2.Fill(ds2, "Siparişler");
            this.dataGridView1.DataSource = ds2.Tables[0];
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form8 frm = new Form8();
            this.Hide();
            frm.yetki = yetki;
            frm.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void Form8_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Hide();
            Form1 form1 = Application.OpenForms["Form1"] as Form1;
            if (form1 != null)
            {
                form1.Show();
            }
        }

            private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            aynısipnugetirme();
        }

        private void dataGridView2_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            aynısipnugetirme();
        }

        private void dataGridView3_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            aynısipnugetirme();
        }

        private void dataGridView4_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            aynısipnugetirme();
        }
        string sipn;
        int satir2;
        int sipnosayısı;
        int x;
        private void dataGridView1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)//farenin sağ tuşuna basılmışsa
            {

                int satir = dataGridView1.HitTest(e.X, e.Y).RowIndex;
                if (satir > -1)
                {
                    dataGridView1.Rows[satir].Selected = true;//bu tıkladığımız alanı seçtiriyoruz
                    sipn = dataGridView1.Rows[satir].Cells["SiparisNo"].Value.ToString();
                }
                satir2 = satir;
            }
        }

        private void dataGridView2_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)//farenin sağ tuşuna basılmışsa
            {

                int satir = dataGridView2.HitTest(e.X, e.Y).RowIndex;
                if (satir > -1)
                {
                    dataGridView2.Rows[satir].Selected = true;//bu tıkladığımız alanı seçtiriyoruz
                    sipn = dataGridView2.Rows[satir].Cells["SiparisNo"].Value.ToString();
                }
                satir2 = satir;
            }
        }

        private void contextMenuStrip1_Click(object sender, EventArgs e)
        {
            sipnosayısı = 0;
            x = 0;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;
            object Missing = Type.Missing;
            //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Enes\\Desktop\\ÜretimFormu.xlsx");
            //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Modeks_Dosyalar\\ÜretimFormu.xlsx ");
            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Open("C:\\Modeks_Dosyalar\\ÜretimFormu.xlsx");

            Microsoft.Office.Interop.Excel.Worksheet sheet2 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT *FROM Siparişler where SiparisNo=@SiparisNo";
            komut.Parameters.AddWithValue("SiparisNo", dataGridView1.Rows[satir2].Cells["SiparisNo"].Value.ToString());
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                sipnosayısı++;
            }
            for (int k = 0; k < sipnosayısı; k++)
            {
                Microsoft.Office.Interop.Excel.Range line = (Microsoft.Office.Interop.Excel.Range)sheet2.Rows[11 + k];
                line.Insert();
            }

            sheet2.Cells[3, 4].Value = sipn; // siparisno yazdırma
            sheet2.Cells[4, 4].Value = dataGridView5.Rows[satir2].Cells["Müşteri"].Value.ToString(); // müşteri yazdırma
            sheet2.Cells[5, 4].Value = dataGridView5.Rows[satir2].Cells["Adres"].Value.ToString(); // adres yazdırma
            sheet2.Cells[7, 4].Value = dataGridView5.Rows[satir2].Cells["Telefon"].Value.ToString(); // telefon yazdırma
            sheet2.Cells[7, 6].Value = dataGridView5.Rows[satir2].Cells["Telefon"].Value.ToString(); // telefon yazdırma
            sheet2.Cells[7, 9].Value = dataGridView5.Rows[satir2].Cells["SiparişTarihi"].Value.ToString(); // sip tarih yazdırma
            sheet2.Cells[8, 9].Value = dataGridView5.Rows[satir2].Cells["TeslimTarihi"].Value.ToString(); // tes tarih yazdırma


            SqlCommand komut2 = new SqlCommand();
            komut2.CommandText = "SELECT *FROM Siparişler where SiparisNo=@SiparisNo";
            komut2.Parameters.AddWithValue("SiparisNo", dataGridView1.Rows[satir2].Cells["SiparisNo"].Value.ToString());
            komut2.Connection = bgl.baglanti();
            komut2.CommandType = CommandType.Text;

            SqlDataReader dr2;
            dr2 = komut2.ExecuteReader();
            while (dr2.Read())
            {
                sheet2.Cells[12 + sipnosayısı, 9].Value = dr2["ToplamM2"].ToString(); // adet toplam yazdırma
                sheet2.Cells[12 + sipnosayısı, 10].Value = dr2["ToplamAdet"].ToString(); // toplam m2 yazdırma
                sheet2.Cells[11 + x, 3].Value = dr2["Model"].ToString(); // model yazdırma
                sheet2.Cells[11 + x, 6].Value = dr2["Özellik"].ToString(); // m2 tarih yazdırma
                sheet2.Cells[11 + x, 4].Value = dr2["Renk"].ToString(); // renk yazdırma
                sheet2.Cells[11 + x, 7].Value = dr2["Boy"].ToString(); // boy yazdırma
                sheet2.Cells[11 + x, 8].Value = dr2["En"].ToString(); // en yazdırma
                sheet2.Cells[11 + x, 9].Value = dr2["ToplamAdet"].ToString(); // adet yazdırma
                sheet2.Cells[11 + x, 10].Value = dr2["M2"].ToString(); // m2 tarih yazdırma
                x++;
            }
            sheet2.PrintPreview();
            workbook.Close(false);
            excel.Quit();
        }

        private void contextMenuStrip2_Click(object sender, EventArgs e)
        {
            sipnosayısı = 0;
            x = 0;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;
            object Missing = Type.Missing;
            //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Enes\\Desktop\\ÜretimFormu.xlsx");
            //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Modeks_Dosyalar\\ÜretimFormu.xlsx ");
            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Open("C:\\Modeks_Dosyalar\\ÜretimFormu.xlsx");

            Microsoft.Office.Interop.Excel.Worksheet sheet2 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT *FROM Siparişler where SiparisNo=@SiparisNo";
            komut.Parameters.AddWithValue("SiparisNo", dataGridView1.Rows[satir2].Cells["SiparisNo"].Value.ToString());
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                sipnosayısı++;
            }
            for (int k = 0; k < sipnosayısı; k++)
            {
                Microsoft.Office.Interop.Excel.Range line = (Microsoft.Office.Interop.Excel.Range)sheet2.Rows[11 + k];
                line.Insert();
            }

            sheet2.Cells[3, 4].Value = sipn; // siparisno yazdırma
            sheet2.Cells[4, 4].Value = dataGridView6.Rows[satir2].Cells["Müşteri"].Value.ToString(); // müşteri yazdırma
            sheet2.Cells[5, 4].Value = dataGridView6.Rows[satir2].Cells["Adres"].Value.ToString(); // adres yazdırma
            sheet2.Cells[7, 4].Value = dataGridView6.Rows[satir2].Cells["Telefon"].Value.ToString(); // telefon yazdırma
            sheet2.Cells[7, 6].Value = dataGridView6.Rows[satir2].Cells["Telefon"].Value.ToString(); // telefon yazdırma
            sheet2.Cells[7, 9].Value = dataGridView6.Rows[satir2].Cells["SiparişTarihi"].Value.ToString(); // sip tarih yazdırma
            sheet2.Cells[8, 9].Value = dataGridView6.Rows[satir2].Cells["TeslimTarihi"].Value.ToString(); // tes tarih yazdırma


            SqlCommand komut2 = new SqlCommand();
            komut2.CommandText = "SELECT *FROM Siparişler where SiparisNo=@SiparisNo";
            komut2.Parameters.AddWithValue("SiparisNo", dataGridView2.Rows[satir2].Cells["SiparisNo"].Value.ToString());
            komut2.Connection = bgl.baglanti();
            komut2.CommandType = CommandType.Text;

            SqlDataReader dr2;
            dr2 = komut2.ExecuteReader();
            while (dr2.Read())
            {
                sheet2.Cells[12 + sipnosayısı, 9].Value = dr2["ToplamM2"].ToString(); // adet toplam yazdırma
                sheet2.Cells[12 + sipnosayısı, 10].Value = dr2["ToplamAdet"].ToString(); // toplam m2 yazdırma
                sheet2.Cells[11 + x, 3].Value = dr2["Model"].ToString(); // model yazdırma
                sheet2.Cells[11 + x, 6].Value = dr2["Özellik"].ToString(); // m2 tarih yazdırma
                sheet2.Cells[11 + x, 4].Value = dr2["Renk"].ToString(); // renk yazdırma
                sheet2.Cells[11 + x, 7].Value = dr2["Boy"].ToString(); // boy yazdırma
                sheet2.Cells[11 + x, 8].Value = dr2["En"].ToString(); // en yazdırma
                sheet2.Cells[11 + x, 9].Value = dr2["ToplamAdet"].ToString(); // adet yazdırma
                sheet2.Cells[11 + x, 10].Value = dr2["M2"].ToString(); // m2 tarih yazdırma
                x++;
            }
            sheet2.PrintPreview();
            workbook.Close(false);
            excel.Quit();
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void button27_Click(object sender, EventArgs e)
        {
            Form18 frm = new Form18();
            frm.yetki = yetki;
            frm.ShowDialog();
        }
        string izin;
        private void MEMBRANDANGERIAL()
        {
            string siparisno = Interaction.InputBox("Membrandan Geri al", "Sipariş Numarasını Giriniz.", "Sipariş No Girin...", 850, 400);

            if (!string.IsNullOrEmpty(siparisno))
            {
                string sorgu = "UPDATE Siparişler SET Aşama=@Aşama,PaketSayısı=null,TeslimEdilenTarih=null,PaketTarihi=null,MembranPressTarihi=null WHERE SiparisNo=@SiparisNo AND TeslimEdilenTarih is null AND PaketTarihi is null AND MembranPressTarihi is null";
                SqlCommand komut;
                komut = new SqlCommand(sorgu, bgl.baglanti());
                komut.Parameters.AddWithValue("@SiparisNo", siparisno);
                komut.Parameters.AddWithValue("@Aşama", "Etiket");
                int etkilenenSatirSayisi = komut.ExecuteNonQuery();

                if (etkilenenSatirSayisi > 0)
                {
                    MessageBox.Show(siparisno + " Sipariş numarası etikete geri gönderilmiştir.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    liste();
                    liste2();
                }
                else
                {
                    MessageBox.Show(siparisno + " Sipariş numarası bulunamadı veya işlem yapılamadı.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                MessageBox.Show("Sipariş numarası girmediniz.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }
        private void button11_Click(object sender, EventArgs e)
        {
            MEMBRANDANGERIAL();
            Methodlar();
            AşağıdakiDatagridMat();
            AşağıdakiDatagridParlak();
        }

        private void comboBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && dataGridView1.Columns[e.ColumnIndex].Name == "SiparisNo")
            {

                string siparisNo = dataGridView1.Rows[e.RowIndex].Cells["SiparisNo"].Value.ToString();
                bool vbkontrol;
                Form10 frm = new Form10();
                frm.siparişno = siparisNo;
                frm.vbkontrol = true;
                frm.yetki = yetki;

                frm.hangiformdan = "Form7";
                this.Hide();
                frm.ShowDialog();
            }

        }

        private void dataGridView2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && dataGridView1.Columns[e.ColumnIndex].Name == "SiparisNo")
            {

                string siparisNo = dataGridView1.Rows[e.RowIndex].Cells["SiparisNo"].Value.ToString();
                bool vbkontrol;
                Form10 frm = new Form10();
                frm.siparişno = siparisNo;
                frm.vbkontrol = true;
                frm.yetki = yetki;

                frm.hangiformdan = "Form7";
                this.Hide();
                frm.ShowDialog();
            }
        }

        private void dataGridView2_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {

        }

        private void button28_Click(object sender, EventArgs e)
        {
            Form2 frm = new Form2();
            frm.yetki = yetki;
            this.Hide();
            frm.Show();
        }
    }
}
