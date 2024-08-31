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
using Microsoft.VisualBasic;
using FirebirdSql.Data.FirebirdClient;
using Microsoft.SqlServer.Management.Smo;
using Microsoft.SqlServer.Management.Common;
using System.IO;

namespace Modeks
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }
        sqlsinif bgl = new sqlsinif();
        public string yetki;
        public string kullaniciadi;
        public string hangiformdan;
        double toplamfiyat = 0;
        double toplamm2 = 0;
        double toplamkapak = 0;
        double kargosipm2 = 0;
        double kargoadet = 0;
        double kargoücret = 0;
        double acilsipm2 = 0;
        double acilsipadet = 0;
        double acilücret = 0;
        double dds = 0;
        double günlüksatıstoplamfiyat = 0;
        double satırsayısı = 0;


        private void müşterigetir()
        {
            comboBox2.Items.Clear();
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (dataGridView1.Rows[i].Visible == true)
                {
                    comboBox2.Items.Add(dataGridView1.Rows[i].Cells["Müşteri"].Value.ToString());
                    satırsayısı++;
                }
            }

            //comboBox2.Items.Clear();
            //SqlCommand komut = new SqlCommand();
            //komut.CommandText = "SELECT DISTINCT Müşteri FROM Siparişler";
            //komut.Connection = bgl.baglanti();
            //komut.CommandType = CommandType.Text;

            //SqlDataReader dr;
            //dr = komut.ExecuteReader();
            //" (dr.Read())
            //{
            //    comboBox2.Items.Add(dr["Müşteri"]);
            //}
        }
        private void müşterigetir_yazarak()
        {
            string srg = comboBox2.Text;
            string sorgu = "SELECT SiparisNo, Onay, Müşteri, Model, Renk, ToplamM2, ToplamAdet, ToplamFiyat, Kargo, AcilFarkı, M2KapakFarkı, ToplamTasarımÜcreti, İskonto, SiparişTarihi, TeslimTarihi, KesildiTarihi, SiparişTipi, SevkTürü, M2KapakAdet, Etiket, MembranPressTarihi, PaketTarihi, TeslimEdilenTarih, PaketSayısı, DDS From Siparişler where Müşteri Like '%" + srg + "%' AND AnaSiparişMi = 'Evet' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
        }
        private void siparişgetir_yazarak()
        {
            string srg = comboBox1.Text;
            string sorgu = "SELECT SiparisNo, Onay, Müşteri, Model, Renk, ToplamM2, ToplamAdet, ToplamFiyat, Kargo, AcilFarkı, M2KapakFarkı, ToplamTasarımÜcreti, İskonto, SiparişTarihi, TeslimTarihi, KesildiTarihi, SiparişTipi, SevkTürü, M2KapakAdet, Etiket, MembranPressTarihi, PaketTarihi, TeslimEdilenTarih, PaketSayısı, DDS, Aşama From Siparişler where SiparisNo Like '%" + srg + "%' AND AnaSiparişMi = 'Evet' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
        }
        private void liste()
        {
            string kayit = "SELECT SiparisNo,Onay,Müşteri,Model,Renk,ToplamM2,ToplamAdet,ToplamFiyat,Kargo,AcilFarkı,M2KapakFarkı,ToplamTasarımÜcreti,İskonto,SiparişTarihi,OnayTarihi,TeslimTarihi,KesildiTarihi,SiparişTipi,SevkTürü,M2KapakAdet,Etiket,MembranPressTarihi,PaketTarihi,TeslimEdilenTarih,PaketSayısı,DDS,Nott,BID, Aşama From Siparişler where AnaSiparişMi=@p1 ORDER BY SiparisNo DESC";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            komut.Parameters.AddWithValue("@p1", "Evet");
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            onayrenklendir();
            aynısipnugetirme();

        }
        private void listedatagrid2()
        {
            string kayit = "SELECT SiparisNo,Onay,Müşteri,Model,Renk,ToplamM2,ToplamAdet,ToplamFiyat,Kargo,AcilFarkı,M2KapakFarkı,ToplamTasarımÜcreti,İskonto,SiparişTarihi,OnayTarihi,TeslimTarihi,KesildiTarihi,SiparişTipi,SevkTürü,M2KapakAdet,Etiket,MembranPressTarihi,PaketTarihi,TeslimEdilenTarih,PaketSayısı,DDS,Nott,BID From Siparişler where AnaSiparişMi=@p1 AND Onay='xxx' ORDER BY SiparisNo DESC";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            komut.Parameters.AddWithValue("@p1", "Evet");
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView2.DataSource = dt;
        }
        private void ToplamFiyat()
        {
            try
            {
                toplamfiyat = 0;
                for (int i = 0; i < dataGridView1.Rows.Count; ++i)
                {
                    if (dataGridView1.Rows[i].Visible == true)
                        toplamfiyat += Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value);
                }
                textBox3.Text = toplamfiyat.ToString("0.##");
            }
            catch (Exception)
            {

            }

        }
        private void ToplamM2()
        {
            try
            {
                toplamm2 = 0;
                for (int i = 0; i < dataGridView1.Rows.Count; ++i)
                {
                    if (dataGridView1.Rows[i].Visible == true)
                        toplamm2 += Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value);
                }
                textBox4.Text = toplamm2.ToString("0.##");
            }
            catch (Exception)
            {

            }

        }
        private void ToplamKapak()
        {
            try
            {
                toplamkapak = 0;
                for (int i = 0; i < dataGridView1.Rows.Count; ++i)
                {
                    if (dataGridView1.Rows[i].Visible == true)
                        toplamkapak += Convert.ToDouble(dataGridView1.Rows[i].Cells[6].Value);
                }
                textBox5.Text = toplamkapak.ToString("0.##");
            }
            catch (Exception)
            {

            }

        }
        private void KargoSipM2()
        {
            try
            {
                kargosipm2 = 0;
                for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
                {
                    if (dataGridView1.Rows[i].Visible == true)
                        if (dataGridView1.Rows[i].Cells[17].Value.ToString() == "Kargo")
                            kargosipm2 += Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value);
                }
                textBox7.Text = kargosipm2.ToString("0.##");
            }
            catch (Exception)
            {

            }

        }
        private void KargoAdet()
        {
            try
            {
                kargoadet = 0;
                for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
                {
                    if (dataGridView1.Rows[i].Visible == true)
                        if (dataGridView1.Rows[i].Cells[17].Value.ToString() == "Kargo")
                            kargoadet += Convert.ToDouble(dataGridView1.Rows[i].Cells[6].Value);
                }
                textBox8.Text = kargoadet.ToString("0.##");
            }
            catch (Exception)
            {

            }

        }
        private void KargoÜcret()
        {
            try
            {
                kargoücret = 0;
                for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
                {
                    if (dataGridView1.Rows[i].Visible == true)
                        kargoücret += Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value);
                }
                textBox9.Text = kargoücret.ToString("0.##");
            }
            catch (Exception)
            {

            }

        }
        private void AcilSipM2()
        {
            try
            {
                acilsipm2 = 0;
                for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
                {
                    if (dataGridView1.Rows[i].Visible == true)
                        if (dataGridView1.Rows[i].Cells[16].Value.ToString() == "Acil")
                            acilsipm2 += Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value);
                }
                textBox10.Text = acilsipm2.ToString("0.##");
            }
            catch (Exception)
            {

            }

        }
        private void AcilSipAdet()
        {
            try
            {
                acilsipadet = 0;
                for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
                {
                    if (dataGridView1.Rows[i].Visible == true)
                        if (dataGridView1.Rows[i].Cells[16].Value.ToString() == "Acil")
                            acilsipadet += Convert.ToDouble(dataGridView1.Rows[i].Cells[6].Value);
                }
                textBox11.Text = acilsipadet.ToString("0.##");
            }
            catch (Exception)
            {

            }

        }
        private void AcilÜcret()
        {
            try
            {
                acilücret = 0;
                for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
                {
                    if (dataGridView1.Rows[i].Visible == true)
                        if (dataGridView1.Rows[i].Cells[16].Value.ToString() == "Acil")
                            acilücret += Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value);
                }
                textBox12.Text = acilücret.ToString("0.##");
            }
            catch (Exception)
            {

            }

        }
        private void DDS()
        {
            try
            {
                dds = 0;
                for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
                {
                    if (dataGridView1.Rows[i].Visible == true)
                        dds += Convert.ToDouble(dataGridView1.Rows[i].Cells["DDS"].Value);
                }
                textBox13.Text = dds.ToString("0.##");
            }
            catch (Exception)
            {

            }

        }
        private void GünlükSatısToplamFiyat()
        {
            try
            {
                textBox1.Visible = true;
                label23.Visible = true;

                DateTime basla = dateTimePicker1.Value.Date;
                DateTime bitir = dateTimePicker2.Value;

                günlüksatıstoplamfiyat = 0;

                HashSet<string> siparisNumaralari = new HashSet<string>(); // Sipariş numaralarını takip etmek için HashSet

                for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
                {
                    if (dataGridView1.Rows[i].Visible == true)
                    {
                        DateTime siparisTarihi = Convert.ToDateTime(dataGridView1.Rows[i].Cells["SiparişTarihi"].Value.ToString());
                        string siparisNo = dataGridView1.Rows[i].Cells["SiparisNo"].Value.ToString();

                        // Sipariş tarihi datetimepicker1 ve datetimepicker2 aralığında mı?
                        if (siparisTarihi >= basla && siparisTarihi <= bitir &&
                            dataGridView1.Rows[i].Cells["Onay"].Value.ToString() == "Onaylandı" &&
                            !siparisNumaralari.Contains(siparisNo)) // Kontrol: Bu sipariş numarası daha önce eklenmedi mi?
                        {
                            günlüksatıstoplamfiyat += Convert.ToDouble(dataGridView1.Rows[i].Cells["ToplamFiyat"].Value);
                            siparisNumaralari.Add(siparisNo); // HashSet'e ekle
                        }
                    }
                }

                textBox1.Text = günlüksatıstoplamfiyat.ToString("0.##");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu: " + ex.Message);
            }


        }
        private void siparisnogetir()
        {
            comboBox1.Items.Clear();
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (dataGridView1.Rows[i].Visible == true)
                {
                    comboBox1.Items.Add(dataGridView1.Rows[i].Cells["SiparisNo"].Value.ToString());
                }
            }

            //SqlCommand komut = new SqlCommand();
            //komut.CommandText = "SELECT DISTINCT SiparisNo FROM Siparişler";
            //komut.Connection = bgl.baglanti();
            //komut.CommandType = CommandType.Text;

            //SqlDataReader dr;
            //dr = komut.ExecuteReader();
            //while (dr.Read())
            //{
            //    comboBox1.Items.Add(dr["SiparisNo"]);
            //}
        }
        private void bidgetir()
        {
            comboBox3.Items.Clear();
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (dataGridView1.Rows[i].Visible == true)
                {
                    comboBox3.Items.Add(dataGridView1.Rows[i].Cells["BID"].Value.ToString());
                }
            }
        }
        private void siparisnoyagöresırala()
        {
            string srg = comboBox1.Text;
            string sorgu = "SELECT SiparisNo,Onay,Müşteri,Model,Renk,ToplamM2,ToplamAdet,ToplamFiyat,Kargo,AcilFarkı,M2KapakFarkı,ToplamTasarımÜcreti,İskonto,SiparişTarihi,OnayTarihi,TeslimTarihi,KesildiTarihi,SiparişTipi,SevkTürü,M2KapakAdet,Etiket,MembranPressTarihi,PaketTarihi,TeslimEdilenTarih,PaketSayısı,DDS,Nott,BID From Siparişler where SiparisNo Like '" + srg + "%' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
        }
        private void bidegöresırala()
        {
            string srg = comboBox3.Text;
            string sorgu = "SELECT SiparisNo,Onay,Müşteri,Model,Renk,ToplamM2,ToplamAdet,ToplamFiyat,Kargo,AcilFarkı,M2KapakFarkı,ToplamTasarımÜcreti,İskonto,SiparişTarihi,OnayTarihi,TeslimTarihi,KesildiTarihi,SiparişTipi,SevkTürü,M2KapakAdet,Etiket,MembranPressTarihi,PaketTarihi,TeslimEdilenTarih,PaketSayısı,DDS,Nott,BID From Siparişler where BID Like '" + srg + "%' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
        }
        private void müsteriyegöresırala()
        {
            string srg = comboBox2.Text;
            string sorgu = "SELECT SiparisNo,Onay,Müşteri,Model,Renk,ToplamM2,ToplamAdet,ToplamFiyat,Kargo,AcilFarkı,M2KapakFarkı,ToplamTasarımÜcreti,İskonto,SiparişTarihi,OnayTarihi,TeslimTarihi,KesildiTarihi,SiparişTipi,SevkTürü,M2KapakAdet,Etiket,MembranPressTarihi,PaketTarihi,TeslimEdilenTarih,PaketSayısı,DDS,Nott,BID From Siparişler where Müşteri Like '" + srg + "%' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
        }
        private void notagöresırala()
        {
            string srg = textBox18.Text;
            string sorgu = "SELECT SiparisNo,Onay,Müşteri,Model,Renk,ToplamM2,ToplamAdet,ToplamFiyat,Kargo,AcilFarkı,M2KapakFarkı,ToplamTasarımÜcreti,İskonto,SiparişTarihi,OnayTarihi,TeslimTarihi,KesildiTarihi,SiparişTipi,SevkTürü,M2KapakAdet,Etiket,MembranPressTarihi,PaketTarihi,TeslimEdilenTarih,PaketSayısı,DDS,Nott,BID From Siparişler where Nott Like '" + srg + "%' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
        }
        private void bugünegöresırala()
        {
            DateTime bitir = DateTime.Now;
            DateTime basla = DateTime.Now;
            dateTimePicker1.Value = basla;
            label21.Text = basla.ToString("yyyy - MM - dd");
            label22.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");

            string sorgu = "SELECT SiparisNo,Onay,Müşteri,Model,Renk,ToplamM2,ToplamAdet,ToplamFiyat,Kargo,AcilFarkı,M2KapakFarkı,ToplamTasarımÜcreti,İskonto,SiparişTarihi,OnayTarihi,TeslimTarihi,KesildiTarihi,SiparişTipi,SevkTürü,M2KapakAdet,Etiket,MembranPressTarihi,PaketTarihi,TeslimEdilenTarih,PaketSayısı,DDS,Nott,BID From Siparişler where SiparişTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
        }
        private void haftayagöresırala()
        {
            DateTime bitir = DateTime.Now;
            DateTime bugun = DateTime.Today;
            DateTime basla = bugun.AddDays(-(int)bugun.DayOfWeek + 1);
            dateTimePicker1.Value = basla;
            label21.Text = basla.ToString("yyyy - MM - dd HH:mm:ss");
            label22.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");

            string sorgu = "SELECT SiparisNo,Onay,Müşteri,Model,Renk,ToplamM2,ToplamAdet,ToplamFiyat,Kargo,AcilFarkı,M2KapakFarkı,ToplamTasarımÜcreti,İskonto,SiparişTarihi,OnayTarihi,TeslimTarihi,KesildiTarihi,SiparişTipi,SevkTürü,M2KapakAdet,Etiket,MembranPressTarihi,PaketTarihi,TeslimEdilenTarih,PaketSayısı,DDS,Nott,BID From Siparişler where SiparişTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
        }
        private void ayagöresırala()
        {
            DateTime bitir = DateTime.Now;
            DateTime bugun = DateTime.Today;
            DateTime basla = new DateTime(bugun.Year, bugun.Month, 1);
            dateTimePicker1.Value = basla;
            label21.Text = basla.ToString("yyyy - MM - dd HH:mm:ss");
            label22.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");
            string sorgu = "SELECT SiparisNo,Onay,Müşteri,Model,Renk,ToplamM2,ToplamAdet,ToplamFiyat,Kargo,AcilFarkı,M2KapakFarkı,ToplamTasarımÜcreti,İskonto,SiparişTarihi,OnayTarihi,TeslimTarihi,KesildiTarihi,SiparişTipi,SevkTürü,M2KapakAdet,Etiket,MembranPressTarihi,PaketTarihi,TeslimEdilenTarih,PaketSayısı,DDS,Nott,BID From Siparişler where SiparişTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];

        }
        private void yılagöresırala()
        {
            DateTime bitir = DateTime.Now;
            DateTime bugun = DateTime.Today;
            DateTime basla = new DateTime(bugun.Year, 1, 1);
            dateTimePicker1.Value = basla;
            label21.Text = basla.ToString("yyyy - MM - dd HH:mm:ss");
            label22.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");

            string sorgu = "SELECT SiparisNo,Onay,Müşteri,Model,Renk,ToplamM2,ToplamAdet,ToplamFiyat,Kargo,AcilFarkı,M2KapakFarkı,ToplamTasarımÜcreti,İskonto,SiparişTarihi,OnayTarihi,TeslimTarihi,KesildiTarihi,SiparişTipi,SevkTürü,M2KapakAdet,Etiket,MembranPressTarihi,PaketTarihi,TeslimEdilenTarih,PaketSayısı,DDS,Nott,BID From Siparişler where SiparişTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];

        }

        private void onaylanmışsiparişler()
        {
            string srg = "Onaylandı";
            string sorgu = "SELECT SiparisNo,Onay,Müşteri,Model,Renk,ToplamM2,ToplamAdet,ToplamFiyat,Kargo,AcilFarkı,M2KapakFarkı,ToplamTasarımÜcreti,İskonto,SiparişTarihi,OnayTarihi,TeslimTarihi,KesildiTarihi,SiparişTipi,SevkTürü,M2KapakAdet,Etiket,MembranPressTarihi,PaketTarihi,TeslimEdilenTarih,PaketSayısı,DDS,Nott,BID From Siparişler where Onay Like '" + srg + "' AND OnayTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
        }
        private void onaybekleyensiparişler()
        {
            string srg = "Onay Bekliyor";
            string sorgu = "SELECT SiparisNo,Onay,Müşteri,Model,Renk,ToplamM2,ToplamAdet,ToplamFiyat,Kargo,AcilFarkı,M2KapakFarkı,ToplamTasarımÜcreti,İskonto,SiparişTarihi,OnayTarihi,TeslimTarihi,KesildiTarihi,SiparişTipi,SevkTürü,M2KapakAdet,Etiket,MembranPressTarihi,PaketTarihi,TeslimEdilenTarih,PaketSayısı,DDS,Nott,BID From Siparişler where Onay Like '" + srg + "' AND SiparişTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
        }
        public void teslimehazırsiparişler()
        {
            string srg = "Hazır";
            string sorgu = "SELECT SiparisNo,Onay,Müşteri,Model,Renk,ToplamM2,ToplamAdet,ToplamFiyat,Kargo,AcilFarkı,M2KapakFarkı,ToplamTasarımÜcreti,İskonto,SiparişTarihi,OnayTarihi,TeslimTarihi,KesildiTarihi,SiparişTipi,SevkTürü,M2KapakAdet,Etiket,MembranPressTarihi,PaketTarihi,TeslimEdilenTarih,PaketSayısı,DDS,Aşama,Nott,BID From Siparişler where PaketTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND  Aşama Like '" + srg + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
        }
        private void acilsiparişler()
        {
            string srg = "Acil";
            string sorgu = "SELECT SiparisNo,Onay,Müşteri,Model,Renk,ToplamM2,ToplamAdet,ToplamFiyat,Kargo,AcilFarkı,M2KapakFarkı,ToplamTasarımÜcreti,İskonto,SiparişTarihi,OnayTarihi,TeslimTarihi,KesildiTarihi,SiparişTipi,SevkTürü,M2KapakAdet,Etiket,MembranPressTarihi,PaketTarihi,TeslimEdilenTarih,PaketSayısı,DDS,Nott,BID From Siparişler where SiparişTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND SiparişTipi Like '" + srg + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
        }
        private void teslimedilensiparişler()
        {
            string srg = "Teslim Edildi";
            string sorgu = "SELECT SiparisNo,Onay,Müşteri,Model,Renk,ToplamM2,ToplamAdet,ToplamFiyat,Kargo,AcilFarkı,M2KapakFarkı,ToplamTasarımÜcreti,İskonto,SiparişTarihi,OnayTarihi,TeslimTarihi,KesildiTarihi,SiparişTipi,SevkTürü,M2KapakAdet,Etiket,MembranPressTarihi,PaketTarihi,TeslimEdilenTarih,PaketSayısı,DDS,Aşama,Nott,BID From Siparişler where TeslimEdilenTarih between '" + label21.Text + "' AND '" + label22.Text + "' AND Aşama Like '" + srg + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
        }
        private void kargosiparişler()
        {
            string srg = "Kargo";
            string sorgu = "SELECT SiparisNo,Onay,Müşteri,Model,Renk,ToplamM2,ToplamAdet,ToplamFiyat,Kargo,AcilFarkı,M2KapakFarkı,ToplamTasarımÜcreti,İskonto,SiparişTarihi,OnayTarihi,TeslimTarihi,KesildiTarihi,SiparişTipi,SevkTürü,M2KapakAdet,Etiket,MembranPressTarihi,PaketTarihi,TeslimEdilenTarih,PaketSayısı,DDS,Nott,BID From Siparişler where TeslimEdilenTarih between '" + label21.Text + "' AND '" + label22.Text + "' AND SevkTürü Like '" + srg + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
        }
        private void İskontoSiparişler()
        {
            string sorgu = "SELECT SiparisNo, Onay, Müşteri, Model, Renk, ToplamM2, ToplamAdet, ToplamFiyat, Kargo, AcilFarkı, M2KapakFarkı, ToplamTasarımÜcreti, İskonto, SiparişTarihi, OnayTarihi, TeslimTarihi, KesildiTarihi, SiparişTipi, SevkTürü, M2KapakAdet, Etiket, MembranPressTarihi, PaketTarihi, TeslimEdilenTarih, PaketSayısı, DDS, Nott, BID FROM Siparişler WHERE SiparişTarihi BETWEEN @StartDate AND @EndDate AND AnaSiparişMi = 'Evet' AND TRY_CONVERT(FLOAT, REPLACE(İskonto, ',', '.')) > 0";
            SqlCommand cmd = new SqlCommand(sorgu, bgl.baglanti());
            cmd.Parameters.AddWithValue("@StartDate", label21.Text);
            cmd.Parameters.AddWithValue("@EndDate", label22.Text);
            SqlDataAdapter adap = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];

        }
        private void parlaksiparişler()
        {
            string srg = "HG%";
            string sorgu = "SELECT SiparisNo,Onay,Müşteri,Model,Renk,ToplamM2,ToplamAdet,ToplamFiyat,Kargo,AcilFarkı,M2KapakFarkı,ToplamTasarımÜcreti,İskonto,SiparişTarihi,OnayTarihi,TeslimTarihi,KesildiTarihi,SiparişTipi,SevkTürü,M2KapakAdet,Etiket,MembranPressTarihi,PaketTarihi,TeslimEdilenTarih,PaketSayısı,DDS,Nott,BID From Siparişler where SiparişTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND Renk Like '" + srg + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
        }
        private void matsiparişler()
        {
            string srg = "HG%";
            string sorgu = "SELECT SiparisNo,Onay,Müşteri,Model,Renk,ToplamM2,ToplamAdet,ToplamFiyat,Kargo,AcilFarkı,M2KapakFarkı,ToplamTasarımÜcreti,İskonto,SiparişTarihi,OnayTarihi,TeslimTarihi,KesildiTarihi,SiparişTipi,SevkTürü,M2KapakAdet,Etiket,MembranPressTarihi,PaketTarihi,TeslimEdilenTarih,PaketSayısı,DDS,Nott,BID From Siparişler where SiparişTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND Renk NOT Like '" + srg + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
        }
        string kesildimi;
        string tarihcnc;
        double bugünkesilencncm2 = 0;
        private void BugünKesilenlerCncM2()
        {
            try
            {
                bugünkesilencncm2 = 0;
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Siparişler where KesildiTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND Onay=@Onay AND AnaSiparişMi=@p1 AND KesildiMi=@KesildiMi";
                komut.Parameters.AddWithValue("@Onay", "Onaylandı");
                komut.Parameters.AddWithValue("@p1", "Evet");
                komut.Parameters.AddWithValue("@KesildiMi", "Evet");
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    kesildimi = dr["KesildiMi"].ToString();
                    tarihcnc = dr["KesildiTarihi"].ToString();

                    if (kesildimi == "Evet")
                    {
                        bugünkesilencncm2 += Convert.ToDouble(dr["ToplamM2"]);
                    }
                }
                textBox14.Text = bugünkesilencncm2.ToString("0.##");
            }
            catch (Exception)
            {

            }

        }
        string tarihetiket;
        double bugünbasılanetiketm2 = 0;
        private void BugünBasılanlarEtiketM2()
        {
            try
            {
                bugünbasılanetiketm2 = 0;
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Siparişler where Etiket between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi=@AnaSiparişMi";
                komut.Parameters.AddWithValue("@AnaSiparişMi", "Evet");
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    tarihetiket = dr["Etiket"].ToString();

                    bugünbasılanetiketm2 += Convert.ToDouble(dr["ToplamM2"]);
                }
                textBox2.Text = bugünbasılanetiketm2.ToString("0.##");
            }
            catch (Exception)
            {

            }

        }
        string tarihmembran;
        double bugünbasılanmembranm2 = 0;
        private void BugünBasılanlarMembranM2()
        {
            try
            {
                bugünbasılanmembranm2 = 0;
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Siparişler where MembranPressTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND  AnaSiparişMi=@AnaSiparişMi";
                komut.Parameters.AddWithValue("@AnaSiparişMi", "Evet");
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    tarihmembran = dr["MembranPressTarihi"].ToString();

                    bugünbasılanmembranm2 += Convert.ToDouble(dr["ToplamM2"]);
                }
                textBox15.Text = bugünbasılanmembranm2.ToString("0.##");
            }
            catch (Exception)
            {

            }

        }
        string tarihpaket;
        double bugünbasılanlarpaketm2 = 0;
        private void BugünBasılanlarPaketM2()
        {
            try
            {
                bugünbasılanlarpaketm2 = 0;
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Siparişler where PaketTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND  AnaSiparişMi=@AnaSiparişMi";
                komut.Parameters.AddWithValue("@AnaSiparişMi", "Evet");
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    tarihpaket = dr["PaketTarihi"].ToString();

                    bugünbasılanlarpaketm2 += Convert.ToDouble(dr["ToplamM2"]);
                }
                textBox16.Text = bugünbasılanlarpaketm2.ToString("0.##");
            }
            catch (Exception)
            {

            }

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
        private void onayrenklendir()
        {
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (Convert.ToString(dataGridView1.Rows[i].Cells[1].Value.ToString()) == "Onaylandı")
                {
                    dataGridView1.Rows[i].Cells[1].Style.BackColor = Color.SeaGreen;
                    dataGridView1.Rows[i].Cells[1].Style.ForeColor = Color.White;
                }
                else if (Convert.ToString(dataGridView1.Rows[i].Cells[1].Value.ToString()) == "Onay Bekliyor")
                {
                    dataGridView1.Rows[i].Cells[1].Style.BackColor = Color.Goldenrod;
                    dataGridView1.Rows[i].Cells[1].Style.ForeColor = Color.White;
                }

                if (Convert.ToString(dataGridView1.Rows[i].Cells["SiparişTipi"].Value.ToString()) == "Acil")
                {
                    dataGridView1.Rows[i].Cells["SiparişTipi"].Style.BackColor = Color.Red;
                    dataGridView1.Rows[i].Cells["SiparişTipi"].Style.ForeColor = Color.White;
                }
                if (Convert.ToString(dataGridView1.Rows[i].Cells["SevkTürü"].Value.ToString()) == "Kargo")
                {
                    dataGridView1.Rows[i].Cells["SevkTürü"].Style.BackColor = Color.CadetBlue;
                    dataGridView1.Rows[i].Cells["SevkTürü"].Style.ForeColor = Color.White;
                }

            }
        }
        public void aynısipnugetirme()
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
        double satırlarınM2Toplamı = 0;
        double satırlarınAdetToplamı = 0;
        double satırlarınFiyatToplamı = 0;
        double satırlarınKargoToplamı = 0;
        double satırlarınAcilFarkToplamı = 0;
        double satırlarınM2KapakFarkToplamı = 0;
        double satırlarınTasarımÜcretiToplamı = 0;
        double satırlarınIskontoToplamı = 0;
        double satırlarınM2KapakAdetToplamı = 0;
        double satırlarınDDSToplamı = 0;
        int birkezcalissin = 0;
        private void SatırlarınEnAltınaToplat()
        {
            try
            {
                satırsayısı = 0;
                müşterigetir();
                satırlarınM2Toplamı = 0;
                satırlarınAdetToplamı = 0;
                satırlarınFiyatToplamı = 0;
                satırlarınKargoToplamı = 0;
                satırlarınAcilFarkToplamı = 0;
                satırlarınM2KapakFarkToplamı = 0;
                satırlarınTasarımÜcretiToplamı = 0;
                satırlarınIskontoToplamı = 0;
                satırlarınM2KapakAdetToplamı = 0;
                satırlarınDDSToplamı = 0;
                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    if (dataGridView1.Rows[i].Visible == true)
                    {
                        satırsayısı++;

                        satırlarınM2Toplamı += Convert.ToDouble(dataGridView1.Rows[i].Cells["ToplamM2"].Value);
                        satırlarınAdetToplamı += Convert.ToDouble(dataGridView1.Rows[i].Cells["ToplamAdet"].Value);
                        satırlarınFiyatToplamı += Convert.ToDouble(dataGridView1.Rows[i].Cells["ToplamFiyat"].Value);
                        satırlarınKargoToplamı += Convert.ToDouble(dataGridView1.Rows[i].Cells["Kargo"].Value);
                        satırlarınAcilFarkToplamı += Convert.ToDouble(dataGridView1.Rows[i].Cells["AcilFarkı"].Value);
                        satırlarınM2KapakFarkToplamı += Convert.ToDouble(dataGridView1.Rows[i].Cells["M2KapakFarkı"].Value);
                        satırlarınTasarımÜcretiToplamı += Convert.ToDouble(dataGridView1.Rows[i].Cells["ToplamTasarımÜcreti"].Value);
                        satırlarınIskontoToplamı += Convert.ToDouble(dataGridView1.Rows[i].Cells["İskonto"].Value);
                        satırlarınM2KapakAdetToplamı += Convert.ToDouble(dataGridView1.Rows[i].Cells["M2KapakAdet"].Value);
                        satırlarınDDSToplamı += Convert.ToDouble(dataGridView1.Rows[i].Cells["DDS"].Value);
                    }
                }
                if (satırsayısı > 0)
                {
                    satırlarınM2KapakAdetToplamı = satırlarınM2KapakAdetToplamı / satırsayısı;
                    dataGridView2.Rows[0].Cells["ToplamM2"].Value = satırlarınM2Toplamı.ToString("#,##0.00");
                    dataGridView2.Rows[0].Cells["ToplamAdet"].Value = satırlarınAdetToplamı.ToString("#,##0.00");
                    dataGridView2.Rows[0].Cells["ToplamFiyat"].Value = satırlarınFiyatToplamı.ToString("#,##0.00");
                    dataGridView2.Rows[0].Cells["Kargo"].Value = satırlarınKargoToplamı.ToString("#,##0.00");
                    dataGridView2.Rows[0].Cells["AcilFarkı"].Value = satırlarınAcilFarkToplamı.ToString("#,##0.00");
                    dataGridView2.Rows[0].Cells["M2KapakFarkı"].Value = satırlarınM2KapakFarkToplamı.ToString("#,##0.00");
                    dataGridView2.Rows[0].Cells["ToplamTasarımÜcreti"].Value = satırlarınTasarımÜcretiToplamı.ToString("#,##0.00");
                    dataGridView2.Rows[0].Cells["İskonto"].Value = satırlarınIskontoToplamı.ToString("#,##0.00");
                    dataGridView2.Rows[0].Cells["M2KapakAdet"].Value = satırlarınM2KapakAdetToplamı.ToString("#,##0.00");
                    dataGridView2.Rows[0].Cells["DDS"].Value = satırlarınDDSToplamı.ToString("#,##0.00");
                }
                dataGridView2.Rows[0].DefaultCellStyle.Font = new Font("Roboto", 12);
                dataGridView2.Rows[0].DefaultCellStyle.ForeColor = Color.Red;
            }
            catch (Exception)
            {

            }
        }
        private void OnBesGunGecmisSiparisiSil()
        {

            string sql = "DELETE FROM Siparişler WHERE Onay = 'Onay Bekliyor' AND DATEDIFF(DAY, SiparişTarihi, GETDATE()) >= 15";
            SqlCommand command = new SqlCommand(sql, bgl.baglanti());
            int rowsAffected = command.ExecuteNonQuery();
            if (rowsAffected > 0)
            {
                MessageBox.Show($"{rowsAffected} satır 15 gündür onaylanmadığı için silindi.");
                liste();
                ToplamFiyat();
                ToplamM2();
                ToplamKapak();
                KargoSipM2();
                KargoAdet();
                KargoÜcret();
                AcilSipM2();
                AcilSipAdet();
                AcilÜcret();
                DDS();
                SatırlarınEnAltınaToplat();
                müşterigetir();
                siparisnogetir();
                bidgetir();
                onayrenklendir();
            }
        }
        private void Form3_Load(object sender, EventArgs e)
        {
            
            SqlConnection.ClearPool(bgl.baglanti());
            if (yetki == "Sekreter")
            {
                groupBox2.Visible = false;
                button17.Visible = false;
                //pictureBox8.Visible = false;
            }
            timer1.Start();
            DateTime bugun = DateTime.Today; // Bugünün tarihini alıyoruz
            DateTime basla = new DateTime(bugun.Year, bugun.Month, bugun.Day, 0, 0, 0); // Bugünün başlangıcını oluşturuyoruz
            DateTime bitir = DateTime.Now; // Şu anki zamanı alıyoruz
            dateTimePicker1.Value = basla; // DateTimePicker'ın değerini başlangıç tarihine ayarlıyoruz
            label21.Text = basla.ToString("yyyy - MM - dd HH:mm:ss"); // Başlangıç tarihini ekranda gösteriyoruz
            label22.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss"); // Bitiş tarihini ekranda göster
            liste();
            listedatagrid2();
            aynısipnugetirme();
            müşterigetir();
            SatırlarınEnAltınaToplat();
            dataGridView1.ClearSelection();
            onayrenklendir();
            siparisnogetir();
            bidgetir();
            müşterigetir();

            ToplamFiyat();
            ToplamM2();
            ToplamKapak();
            KargoSipM2();
            KargoAdet();
            KargoÜcret();
            AcilSipM2();
            AcilSipAdet();
            AcilÜcret();
            DDS();
            GünlükSatısToplamFiyat();


            BugünKesilenlerCncM2();
            BugünBasılanlarEtiketM2();
            BugünBasılanlarMembranM2();
            BugünBasılanlarPaketM2();
            dataGridView1.Focus();
            OnBesGunGecmisSiparisiSil();
        }

        private void pictureBox8_Click(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            label2.Text = DateTime.Now.ToLongDateString();
            label12.Text = DateTime.Now.ToLongTimeString();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Form10 frm = new Form10();
            frm.hangiformdan = hangiformdan;
            frm.kullaniciadi = kullaniciadi;
            frm.yetki = yetki;
            this.Hide();
            frm.ShowDialog();
        }
        string siparişno;
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {


        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            siparişno = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            label20.Text = siparişno;
        }

        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Form10 frm = new Form10();
            frm.hangiformdan = hangiformdan;
            frm.siparişno = siparişno;
            frm.kullaniciadi = kullaniciadi;
            this.Hide();
            frm.ShowDialog();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            siparisnoyagöresırala();
            aynısipnugetirme();
            SatırlarınEnAltınaToplat();
            onayrenklendir();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            liste();
            ToplamFiyat();
            ToplamM2();
            ToplamKapak();
            KargoSipM2();
            KargoAdet();
            KargoÜcret();
            AcilSipM2();
            AcilSipAdet();
            AcilÜcret();
            DDS();
            SatırlarınEnAltınaToplat();
            müşterigetir();
            siparisnogetir();
            bidgetir();
            onayrenklendir();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            müsteriyegöresırala();
            aynısipnugetirme();
            SatırlarınEnAltınaToplat();
            onayrenklendir();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            bugünegöresırala();
            GünlükSatısToplamFiyat();
            aynısipnugetirme();
            ToplamFiyat();
            ToplamM2();
            ToplamKapak();
            KargoSipM2();
            KargoAdet();
            KargoÜcret();
            AcilSipM2();
            AcilSipAdet();
            AcilÜcret();
            DDS();
            BugünKesilenlerCncM2();
            BugünBasılanlarEtiketM2();
            BugünBasılanlarMembranM2();
            BugünBasılanlarPaketM2();
            SatırlarınEnAltınaToplat();
            onayrenklendir();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            haftayagöresırala();
            GünlükSatısToplamFiyat();
            aynısipnugetirme();
            ToplamFiyat();
            ToplamM2();
            ToplamKapak();
            KargoSipM2();
            KargoAdet();
            KargoÜcret();
            AcilSipM2();
            AcilSipAdet();
            AcilÜcret();
            DDS();
            BugünKesilenlerCncM2();
            BugünBasılanlarEtiketM2();
            BugünBasılanlarMembranM2();
            BugünBasılanlarPaketM2();
            SatırlarınEnAltınaToplat();
            onayrenklendir();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            parlaksiparişler();
            GünlükSatısToplamFiyat();
            aynısipnugetirme();
            ToplamFiyat();
            ToplamM2();
            ToplamKapak();
            KargoSipM2();
            KargoAdet();
            KargoÜcret();
            AcilSipM2();
            AcilSipAdet();
            AcilÜcret();
            DDS();
            SatırlarınEnAltınaToplat();
            müşterigetir();
            siparisnogetir();
            bidgetir();
            onayrenklendir();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            matsiparişler();
            GünlükSatısToplamFiyat();
            aynısipnugetirme();
            ToplamFiyat();
            ToplamM2();
            ToplamKapak();
            KargoSipM2();
            KargoAdet();
            KargoÜcret();
            AcilSipM2();
            AcilSipAdet();
            AcilÜcret();
            DDS();
            SatırlarınEnAltınaToplat();
            müşterigetir();
            siparisnogetir();
            bidgetir();
            onayrenklendir();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            onaylanmışsiparişler();
            GünlükSatısToplamFiyat();
            aynısipnugetirme();
            ToplamFiyat();
            ToplamM2();
            ToplamKapak();
            KargoSipM2();
            KargoAdet();
            KargoÜcret();
            AcilSipM2();
            AcilSipAdet();
            AcilÜcret();
            DDS();
            SatırlarınEnAltınaToplat();
            müşterigetir();
            siparisnogetir();
            bidgetir();
            onayrenklendir();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            onaybekleyensiparişler();
            GünlükSatısToplamFiyat();
            aynısipnugetirme();
            ToplamFiyat();
            ToplamM2();
            ToplamKapak();
            KargoSipM2();
            KargoAdet();
            KargoÜcret();
            AcilSipM2();
            AcilSipAdet();
            AcilÜcret();
            DDS();
            SatırlarınEnAltınaToplat();
            müşterigetir();
            siparisnogetir();
            bidgetir();
            onayrenklendir();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            teslimehazırsiparişler();
            GünlükSatısToplamFiyat();
            aynısipnugetirme();
            ToplamFiyat();
            ToplamM2();
            ToplamKapak();
            KargoSipM2();
            KargoAdet();
            KargoÜcret();
            AcilSipM2();
            AcilSipAdet();
            AcilÜcret();
            DDS();
            SatırlarınEnAltınaToplat();
            müşterigetir();
            siparisnogetir();
            bidgetir();
            onayrenklendir();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            acilsiparişler();
            GünlükSatısToplamFiyat();
            aynısipnugetirme();
            ToplamFiyat();
            ToplamM2();
            ToplamKapak();
            KargoSipM2();
            KargoAdet();
            KargoÜcret();
            AcilSipM2();
            AcilSipAdet();
            AcilÜcret();
            DDS();
            SatırlarınEnAltınaToplat();
            müşterigetir();
            siparisnogetir();
            bidgetir();
            onayrenklendir();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            liste();
            teslimedilensiparişler();
            GünlükSatısToplamFiyat();
            aynısipnugetirme();
            ToplamFiyat();
            ToplamM2();
            ToplamKapak();
            KargoSipM2();
            KargoAdet();
            KargoÜcret();
            AcilSipM2();
            AcilSipAdet();
            AcilÜcret();
            DDS();
            SatırlarınEnAltınaToplat();
            müşterigetir();
            siparisnogetir();
            bidgetir();
            onayrenklendir();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            kargosiparişler();
            GünlükSatısToplamFiyat();
            aynısipnugetirme();
            ToplamFiyat();
            ToplamM2();
            ToplamKapak();
            KargoSipM2();
            KargoAdet();
            KargoÜcret();
            AcilSipM2();
            AcilSipAdet();
            AcilÜcret();
            DDS();
            SatırlarınEnAltınaToplat();
            müşterigetir();
            siparisnogetir();
            bidgetir();
            onayrenklendir();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ayagöresırala();
            GünlükSatısToplamFiyat();
            aynısipnugetirme();
            ToplamFiyat();
            ToplamM2();
            ToplamKapak();
            KargoSipM2();
            KargoAdet();
            KargoÜcret();
            AcilSipM2();
            AcilSipAdet();
            AcilÜcret();
            DDS();
            BugünKesilenlerCncM2();
            BugünBasılanlarEtiketM2();
            BugünBasılanlarMembranM2();
            BugünBasılanlarPaketM2();
            SatırlarınEnAltınaToplat();
            onayrenklendir();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            yılagöresırala();
            GünlükSatısToplamFiyat();
            aynısipnugetirme();
            ToplamFiyat();
            ToplamM2();
            ToplamKapak();
            KargoSipM2();
            KargoAdet();
            KargoÜcret();
            AcilSipM2();
            AcilSipAdet();
            AcilÜcret();
            DDS();
            BugünKesilenlerCncM2();
            BugünBasılanlarEtiketM2();
            BugünBasılanlarMembranM2();
            BugünBasılanlarPaketM2();
            SatırlarınEnAltınaToplat();
            onayrenklendir();
        }

        string izin;
        private void button17_Click(object sender, EventArgs e)
        {
            string siparisno = Interaction.InputBox("Onay Kaldırma", "Onayı kaldırılacak sipariş numarası giriniz.", "Sipariş No Girin...", 850, 400);
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString() == siparisno && dataGridView1.Rows[i].Cells[1].Value.ToString() == "Onaylandı")
                {
                    izin = "var";
                }
            }

            if (izin == "var")
            {
                string sorgu = "UPDATE Siparişler SET Onay=@Onay WHERE SiparisNo=@SiparisNo";
                SqlCommand komut;
                komut = new SqlCommand(sorgu, bgl.baglanti());
                komut.Parameters.AddWithValue("@SiparisNo", siparisno);
                komut.Parameters.AddWithValue("@Onay", "Onay Bekliyor");
                komut.ExecuteNonQuery();
                MessageBox.Show(siparisno + " Sipariş numarasının onayı kaldırılmıştır.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                liste();

                //BSR DEN SİLME
                var connectionString = @"User ID=SYSDBA;Password=masterkey;Database=localhost:C:\Program Files (x86)\BSR\VERESIYEDATA.FDB ;Charset=WIN1254;";
                FbConnection fbcnn = new FbConnection(connectionString);//bağlan
                fbcnn.Open();
                string sorgusil = "DELETE FROM SATISLAR WHERE ACIKLAMA=@ACIKLAMA";
                FbCommand komutsil = new FbCommand(sorgusil, fbcnn);
                komutsil.Parameters.AddWithValue("@ACIKLAMA", siparisno);
                komutsil.ExecuteNonQuery();
                fbcnn.Close();
                MessageBox.Show("Kayıt BSR'Den başarıyla silindi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Lütfen geçerli bir sipariş numarası giriniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }
        //string izin2;
        private void button11_Click(object sender, EventArgs e)
        {
            İskontoSiparişler();
            GünlükSatısToplamFiyat();
            aynısipnugetirme();
            ToplamFiyat();
            ToplamM2();
            ToplamKapak();
            KargoSipM2();
            KargoAdet();
            KargoÜcret();
            AcilSipM2();
            AcilSipAdet();
            AcilÜcret();
            DDS();
            SatırlarınEnAltınaToplat();
            müşterigetir();
            siparisnogetir();
            bidgetir();
            onayrenklendir();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTime bitir = dateTimePicker2.Value;
            DateTime basla = dateTimePicker1.Value;
            label21.Text = basla.ToString("yyyy - MM - dd  HH:mm:ss");
            label22.Text = bitir.ToString("yyyy - MM - dd  HH:mm:ss");

            string sorgu = "SELECT SiparisNo,Onay,Müşteri,Model,Renk,ToplamM2,ToplamAdet,ToplamFiyat,Kargo,AcilFarkı,M2KapakFarkı,ToplamTasarımÜcreti,İskonto,SiparişTarihi,OnayTarihi,TeslimTarihi,KesildiTarihi,SiparişTipi,SevkTürü,M2KapakAdet,Etiket,MembranPressTarihi,PaketTarihi,PaketSayısı,DDS From Siparişler where SiparişTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            GünlükSatısToplamFiyat();
            aynısipnugetirme();
            ToplamFiyat();
            ToplamM2();
            ToplamKapak();
            KargoSipM2();
            KargoAdet();
            KargoÜcret();
            AcilSipM2();
            AcilSipAdet();
            AcilÜcret();
            DDS();
            BugünKesilenlerCncM2();
            BugünBasılanlarEtiketM2();
            BugünBasılanlarMembranM2();
            BugünBasılanlarPaketM2();
            SatırlarınEnAltınaToplat();
            onayrenklendir();
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            DateTime bitir = dateTimePicker2.Value;
            DateTime basla = dateTimePicker1.Value;
            label21.Text = basla.ToString("yyyy - MM - dd  HH:mm:ss");
            label22.Text = bitir.ToString("yyyy - MM - dd  HH:mm:ss");

            string sorgu = "SELECT SiparisNo,Onay,Müşteri,Model,Renk,ToplamM2,ToplamAdet,ToplamFiyat,Kargo,AcilFarkı,M2KapakFarkı,ToplamTasarımÜcreti,İskonto,SiparişTarihi,OnayTarihi,TeslimTarihi,KesildiTarihi,SiparişTipi,SevkTürü,M2KapakAdet,Etiket,MembranPressTarihi,PaketTarihi,TeslimEdilenTarih,PaketSayısı,DDS From Siparişler where SiparişTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            GünlükSatısToplamFiyat();
            aynısipnugetirme();
            ToplamFiyat();
            ToplamM2();
            ToplamKapak();
            KargoSipM2();
            KargoAdet();
            KargoÜcret();
            AcilSipM2();
            AcilSipAdet();
            AcilÜcret();
            DDS();
            BugünKesilenlerCncM2();
            BugünBasılanlarEtiketM2();
            BugünBasılanlarMembranM2();
            BugünBasılanlarPaketM2();
            SatırlarınEnAltınaToplat();
            onayrenklendir();
        }

        private void button18_Click(object sender, EventArgs e)
        {
            Form3 frm = new Form3();
            this.Hide();
            frm.yetki = yetki;
            frm.kullaniciadi = kullaniciadi;
            frm.hangiformdan = hangiformdan;
            frm.Show();
        }

        private void comboBox2_TextChanged(object sender, EventArgs e)
        {
            müşterigetir_yazarak();
            aynısipnugetirme();
            onayrenklendir();
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            siparişgetir_yazarak();
            aynısipnugetirme();
            onayrenklendir();
            AsamaGetir();
        }

        void AsamaGetir()
        {
            label34.Text = "";
            if (dataGridView1.RowCount > 1)
            {
                string asama = dataGridView1.Rows[0].Cells["Aşama"].Value.ToString();
                string TeslimEdilenTarih = dataGridView1.Rows[0].Cells["TeslimEdilenTarih"].Value.ToString();

                if (asama == "Hazır") asama = "Teslime Hazır";
                if (asama == "Onaylandı") asama = "Onaylandı CNC'de";
                else if (asama == "Kargo") asama = "Pakette";
                else if (asama == "Etiket") asama = "CNC Kesildi Etiket Palette";
                else if (asama == "Palet") asama = "Presste";

                if (asama != null) button22.Text = asama;
                else button22.Text = "";

                if (TeslimEdilenTarih != "")
                {
                    label34.Text = "Teslim Edilen Tarih: " + TeslimEdilenTarih;
                }
            }
        }

        private void Form3_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void button19_Click(object sender, EventArgs e)
        {
            Form19 frm = new Form19();
            frm.yetki = yetki;
            if (hangiformdan == "form23")
                frm.hangiformdan = "form23";
            else
                frm.hangiformdan = "Form3";
            this.Hide();
            frm.ShowDialog();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            Form18 frm = new Form18();
            if (hangiformdan == "form23")
                frm.hangiformdan = "form23";
            else
                frm.hangiformdan = "Form3";
            frm.kullaniciadi = kullaniciadi;
            frm.yetki = yetki;
            this.Hide();
            frm.ShowDialog();
        }

        private void DbBackup_Complete(object sender, ServerMessageEventArgs e)
        {
            try
            {
                MessageBox.Show("Yedekleme işlemi başarılı bir şekilde gerçekleşti.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView1_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {

        }

        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            aynısipnugetirme();
            SatırlarınEnAltınaToplat();
            onayrenklendir();
        }

        private void button20_Click(object sender, EventArgs e)
        {
            string database = "Modeks_2022";

            // SqlConnection nesnesi oluşturma
            using (SqlConnection connection = new SqlConnection(@"Data Source=85.187.184.44,1433;Initial Catalog=Modeks_2022;User ID=sa;Password=modeks@123"))
            {
                // Bağlantıyı açma
                connection.Open();

                // SaveFileDialog oluşturma
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Filter = "Backup Files (*.bak)|*.bak";
                saveFileDialog1.Title = "Veritabanı Yedekleme Dosyası Kaydet";
                saveFileDialog1.FileName = DateTime.Now.ToShortDateString() + "backup.bak";
                string filePath = saveFileDialog1.FileName;
                string directoryPath = $"C:\\Yedekler\\{saveFileDialog1.FileName}";

                // Yedekleme işlemi için SQL sorgusu
                string backupQuery = $"BACKUP DATABASE {database} TO DISK='{directoryPath}'";

                // SQL sorgusu için SqlCommand nesnesi oluşturma
                using (SqlCommand command = new SqlCommand(backupQuery, connection))
                {
                    // Yedekleme işlemini gerçekleştirme
                    command.ExecuteNonQuery();

                    MessageBox.Show("Veritabanı yedekleme işlemi başarılı.");

                    //File.Copy(directoryPath, $"C:\\Users\\Administrator\\Dropbox\\{saveFileDialog1.FileName}", true);
                }
                // Bağlantıyı kapatma
                connection.Close();
            }
        }

        private void dataGridView1_CellMouseClick_1(object sender, DataGridViewCellMouseEventArgs e)
        {

        }

        private void dataGridView1_Click(object sender, EventArgs e)
        {
            SatırlarınEnAltınaToplat();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            bidegöresırala();
            aynısipnugetirme();
            SatırlarınEnAltınaToplat();
            onayrenklendir();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void comboBox3_TextChanged(object sender, EventArgs e)
        {
            bidegöresırala();
            aynısipnugetirme();
            SatırlarınEnAltınaToplat();
            onayrenklendir();
        }

        private void button21_Click(object sender, EventArgs e)
        {
            if (hangiformdan == "form23")
            {
                Form23 yeni = new Form23();
                yeni.yetki = yetki;
                yeni.kullaniciadi = kullaniciadi;
                yeni.hangiformdan = hangiformdan;
                this.Hide();
                yeni.Show();

            }
            else
            {
                Form2 frm = new Form2();
                frm.yetki = yetki;
                frm.kullaniciadi = kullaniciadi;
                this.Hide();
                frm.Show();
            }
        }
        private void CopySelectedCellsToClipboard()
        {
            // Seçili hücrelerin içeriğini birleştirerek panoya kopyala
            string cellValues = string.Join("\t", dataGridView1.SelectedCells.Cast<DataGridViewCell>().Select(cell => cell.Value));
            Clipboard.SetText(cellValues);
        }

        private void kopyalaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Seçili hücreleri kopyala
            DataObject dataObject = dataGridView1.GetClipboardContent();
            Clipboard.SetDataObject(dataObject);
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
        }

        private void textBox18_TextChanged(object sender, EventArgs e)
        {
            notagöresırala();
            aynısipnugetirme();
            SatırlarınEnAltınaToplat();
            onayrenklendir();
        }

        private void dataGridView1_Scroll(object sender, ScrollEventArgs e)
        {
            if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll)
            {
                dataGridView2.FirstDisplayedScrollingRowIndex = dataGridView1.FirstDisplayedScrollingRowIndex;
                dataGridView2.HorizontalScrollingOffset = dataGridView1.HorizontalScrollingOffset;
            }
        }

        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label33_Click(object sender, EventArgs e)
        {

        }

        private void button22_Click(object sender, EventArgs e)
        {

        }
    }
}
