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
    public partial class Form6 : Form
    {
        public Form6()
        {
            InitializeComponent();
        }
        sqlsinif bgl = new sqlsinif();
        public string yetki;
        double kesilmişkesilecekm2 = 0;
        double kesilmişkesilecekadet = 0;
        double matm2 = 0;
        double parlakm2 = 0;
        double matadet = 0;
        double parlakadet = 0;
        double acilmatm2 = 0;
        double acilparlakm2 = 0;
        double acilmatadet = 0;
        double acilparlakadet = 0;
        double kesilecekmatparlakm2 = 0;
        double kesilecekmatparlakadet = 0;

        double kesilenmatm2 = 0;
        double kesilenparlakm2 = 0;
        double kesilenmatadet = 0;
        double kesilenparlakadet = 0;
        double kesilenacilmatm2 = 0;
        double kesilenacilparlakm2 = 0;
        double kesilenacilmatadet = 0;
        double kesilenacilparlakadet = 0;
        double kesilentoplamm2 = 0;
        double kesilentoplamadet = 0;

        double kesilenmatm2_ = 0;
        double kesilenparlakm2_ = 0;
        double kesilenmatadet_ = 0;
        double kesilenparlakadet_ = 0;
        double kesilenacilmatm2_ = 0;
        double kesilenacilparlakm2_ = 0;
        double kesilenacilmatadet_ = 0;
        double kesilenacilparlakadet_ = 0;
        double kesilentoplamm2_ = 0;
        double kesilentoplamadet_ = 0;
        double kesilentoplamm2bölüadet_ = 0;

        private void liste()
        {
            string kayit = "SELECT DISTINCT SiparisNo,TeslimTarihi,SiparişTarihi,Müşteri as'Firma Ünvanı',Model,Renk,SiparişTipi,ToplamM2,Onay,KesildiTarihi,ToplamAdet From Siparişler where  Onay=@Onay AND AnaSiparişMi=@p1 AND KesildiMi=@KesildiMi ORDER BY SiparisNo DESC";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            komut.Parameters.AddWithValue("@Onay", "Onaylandı");
            komut.Parameters.AddWithValue("@p1", "Evet");
            komut.Parameters.AddWithValue("@KesildiMi", "Hayır");
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            onayrenklendir();
        }
        private void liste2()
        {
            string kayit = "SELECT DISTINCT SiparisNo,TeslimTarihi,SiparişTarihi,Müşteri as'Firma Ünvanı',Model,Renk,SiparişTipi,ToplamM2,Onay,KesildiTarihi,ToplamAdet,Adres,Telefon From Siparişler where Onay=@Onay AND AnaSiparişMi=@p1 AND KesildiMi=@KesildiMi ORDER BY SiparisNo ASC";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            komut.Parameters.AddWithValue("@Onay", "Onaylandı");
            komut.Parameters.AddWithValue("@p1", "Evet");
            komut.Parameters.AddWithValue("@KesildiMi", "Hayır");
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView2.DataSource = dt;
            onayrenklendir();
        }
        private void onayrenklendir()
        {
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (Convert.ToString(dataGridView1.Rows[i].Cells["Onay"].Value.ToString()) == "Onaylandı")
                {
                    dataGridView1.Rows[i].Cells["Onay"].Style.BackColor = Color.SeaGreen;
                    dataGridView1.Rows[i].Cells["Onay"].Style.ForeColor = Color.White;
                }
                else if (Convert.ToString(dataGridView1.Rows[i].Cells["Onay"].Value.ToString()) == "Onay Bekliyor")
                {
                    dataGridView1.Rows[i].Cells["Onay"].Style.BackColor = Color.Goldenrod;
                    dataGridView1.Rows[i].Cells["Onay"].Style.ForeColor = Color.White;
                }

                if (Convert.ToString(dataGridView1.Rows[i].Cells["SiparişTipi"].Value.ToString()) == "Acil")
                {
                    dataGridView1.Rows[i].Cells["SiparişTipi"].Style.BackColor = Color.Red;
                    dataGridView1.Rows[i].Cells["SiparişTipi"].Style.ForeColor = Color.White;
                }
                else if (Convert.ToString(dataGridView1.Rows[i].Cells["SiparişTipi"].Value.ToString()) == "Normal")
                {
                    dataGridView1.Rows[i].Cells["SiparişTipi"].Style.BackColor = Color.Orange;
                    dataGridView1.Rows[i].Cells["SiparişTipi"].Style.ForeColor = Color.White;
                }

            }
        }
        private void siparisnogetir()
        {
            comboBox1.Items.Clear();
            comboBox3.Items.Clear();
            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT DISTINCT SiparisNo FROM Siparişler Where Onay=@Onay AND AnaSiparişMi=@p1 AND KesildiMi=@KesildiMi";
            komut.Parameters.AddWithValue("@Onay", "Onaylandı");
            komut.Parameters.AddWithValue("@p1", "Evet");
            komut.Parameters.AddWithValue("@KesildiMi", "Hayır");
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                comboBox1.Items.Add(dr["SiparisNo"]);
                comboBox3.Items.Add(dr["SiparisNo"]);
            }
            onayrenklendir();
        }
        private void renkgetir()
        {
            comboBox2.Items.Clear();
            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT DISTINCT Renk FROM Siparişler";
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                comboBox2.Items.Add(dr["Renk"]);
            }
            onayrenklendir();
        }
        private void siparisnoyagöresırala()
        {
            string srg = comboBox1.Text;
            string sorgu = "Select DISTINCT SiparisNo,TeslimTarihi,SiparişTarihi,Müşteri as'Firma Ünvanı',Model,Renk,SiparişTipi,ToplamM2,Onay,KesildiTarihi,ToplamAdet from Siparişler where SiparisNo Like '" + srg + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            onayrenklendir();
        }
        private void rengegöresırala()
        {
            string srg = comboBox2.Text;
            string sorgu = "Select SiparisNo,TeslimTarihi,SiparişTarihi,Müşteri as'Firma Ünvanı',Model,Renk,SiparişTipi,ToplamM2,Onay,KesildiTarihi,ToplamAdet from Siparişler where Renk Like '" + srg + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            onayrenklendir();
        }
        private void KesilmişKesilecekM2()
        {
            kesilmişkesilecekm2 = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true)
                    kesilmişkesilecekm2 += Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value);
            }
            textBox3.Text = kesilmişkesilecekm2.ToString("0.##");
        }
        private void KesilmişKesilecekAdet()
        {
            kesilmişkesilecekadet = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true)
                    kesilmişkesilecekadet += Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value);
            }
            textBox4.Text = kesilmişkesilecekadet.ToString("0.##");
        }
        private void MatParlakM2()
        {
            matm2 = 0;
            parlakm2 = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true)
                    if (dataGridView1.Rows[i].Cells[5].Value.ToString().StartsWith("HG") &&( dataGridView1.Rows[i].Cells[6].Value.ToString() == "Normal" || dataGridView1.Rows[i].Cells[6].Value.ToString() == "Üretim Sorunu"))
                        if (dataGridView1.Rows[i].Cells[9].Value.ToString().Length < 1)
                            parlakm2 += Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value);
                        else
                        {

                        }

                    else if (dataGridView1.Rows[i].Cells[9].Value.ToString().Length < 1 && (dataGridView1.Rows[i].Cells[6].Value.ToString() == "Normal" || dataGridView1.Rows[i].Cells[6].Value.ToString() == "Üretim Sorunu"))
                        matm2 += Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value);
            }
            textBox5.Text = matm2.ToString("0.##");
            textBox8.Text = parlakm2.ToString("0.##");
        }
        private void MatParlakAdet()
        {
            matadet = 0;
            parlakadet = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true)
                    if (dataGridView1.Rows[i].Cells[5].Value.ToString().StartsWith("HG") && (dataGridView1.Rows[i].Cells[6].Value.ToString() == "Normal" || dataGridView1.Rows[i].Cells[6].Value.ToString() == "Üretim Sorunu"))
                        if (dataGridView1.Rows[i].Cells[9].Value.ToString().Length < 1)
                            parlakadet += Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value);
                        else
                        {

                        }
                    else if (dataGridView1.Rows[i].Cells[9].Value.ToString().Length < 1 && (dataGridView1.Rows[i].Cells[6].Value.ToString() == "Normal" || dataGridView1.Rows[i].Cells[6].Value.ToString() == "Üretim Sorunu"))
                        matadet += Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value);
            }
            textBox6.Text = matadet.ToString("0.##");
            textBox7.Text = parlakadet.ToString("0.##");
        }
        private void AcilMatParlakM2()
        {
            acilmatm2 = 0;
            acilparlakm2 = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true)
                    if (dataGridView1.Rows[i].Cells[5].Value.ToString().StartsWith("HG") && dataGridView1.Rows[i].Cells[6].Value.ToString() == "Acil")
                        if (dataGridView1.Rows[i].Cells[9].Value.ToString().Length < 1)
                            acilparlakm2 += Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value);
                        else
                        {

                        }
                    else if (dataGridView1.Rows[i].Cells[6].Value.ToString() == "Acil" && dataGridView1.Rows[i].Cells[9].Value.ToString().Length < 1)
                        acilmatm2 += Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value);
            }
            textBox10.Text = acilmatm2.ToString("0.##");
            textBox12.Text = acilparlakm2.ToString("0.##");
        }
        private void AcilMatParlakAdet()
        {
            acilmatadet = 0;
            acilparlakadet = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true)
                    if (dataGridView1.Rows[i].Cells[5].Value.ToString().StartsWith("HG") && dataGridView1.Rows[i].Cells[6].Value.ToString() == "Acil")
                        if (dataGridView1.Rows[i].Cells[9].Value.ToString().Length < 1)
                            acilparlakadet += Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value);
                        else
                        {

                        }
                    else if (dataGridView1.Rows[i].Cells[6].Value.ToString() == "Acil" && dataGridView1.Rows[i].Cells[9].Value.ToString().Length < 1)
                        acilmatadet += Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value);
            }
            textBox9.Text = acilmatadet.ToString("0.##");
            textBox11.Text = acilparlakadet.ToString("0.##");
        }
        private void KesilecekMatParlakM2()
        {
            kesilecekmatparlakm2 = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true)
                    if (dataGridView1.Rows[i].Cells[9].Value.ToString().Length < 5)
                        kesilecekmatparlakm2 += Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value);
            }
            textBox14.Text = kesilecekmatparlakm2.ToString("0.##");
        }
        private void KesilecekMatParlakAdet()
        {
            kesilecekmatparlakadet = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true)
                    if (dataGridView1.Rows[i].Cells[9].Value.ToString().Length < 5)
                        kesilecekmatparlakadet += Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value);
            }
            textBox13.Text = kesilecekmatparlakadet.ToString("0.##");
        }
        string kesildimi;
        DateTime tarih;
        string renk;
        string sipariştipi;
        private void BugünKesilenlerr()
        {
            kesilenmatm2 = 0;
            kesilenmatadet = 0;
            kesilenparlakm2 = 0;
            kesilenparlakadet = 0;
            kesilenacilmatm2 = 0;
            kesilenacilmatadet = 0;
            kesilenacilparlakm2 = 0;
            kesilenacilparlakadet = 0;
            kesilentoplamm2 = 0;
            kesilentoplamadet = 0;

            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT *FROM Siparişler where Onay=@Onay AND AnaSiparişMi=@p1 AND KesildiMi=@KesildiMi";
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
                tarih = Convert.ToDateTime(dr["KesildiTarihi"].ToString());
                renk = dr["Renk"].ToString();
                sipariştipi = dr["SiparişTipi"].ToString();

                //acil olmayan kısımlar

                if (dr["KesildiTarihi"].ToString() != "")
                {
                    tarih = Convert.ToDateTime(dr["KesildiTarihi"].ToString());
                    if (kesildimi == "Evet" && sipariştipi == "Acil" && tarih.ToString("yyyy - MM - dd") == (DateTime.Now.ToString("yyyy - MM - dd")))
                    {
                        if (renk.StartsWith("HG"))
                        {
                            kesilenacilparlakm2 += Convert.ToDouble(dr["M2"]);
                            kesilenacilparlakadet += Convert.ToDouble(dr["Adet"]);
                            textBox20.Text = kesilenacilparlakm2.ToString("0.##");
                            textBox19.Text = kesilenacilparlakadet.ToString("0.##");
                        }
                        else
                        {
                            kesilenacilmatm2 += Convert.ToDouble(dr["M2"]);
                            kesilenacilmatadet += Convert.ToDouble(dr["Adet"]);
                            textBox22.Text = kesilenacilmatm2.ToString("0.##");
                            textBox21.Text = kesilenacilmatadet.ToString("0.##");
                        }
                    }
                    else if (kesildimi == "Evet" && tarih.ToString("yyyy - MM - dd") == (DateTime.Now.ToString("yyyy - MM - dd")))
                    {
                        if (renk.StartsWith("HG"))
                        {
                            kesilenparlakm2 += Convert.ToDouble(dr["M2"]);
                            kesilenparlakadet += Convert.ToDouble(dr["Adet"]);
                            textBox24.Text = kesilenparlakm2.ToString("0.##");
                            textBox23.Text = kesilenparlakadet.ToString("0.##");
                        }
                        else
                        {
                            kesilenmatm2 += Convert.ToDouble(dr["M2"]);
                            kesilenmatadet += Convert.ToDouble(dr["Adet"]);
                            textBox26.Text = kesilenmatm2.ToString("0.##");
                            textBox25.Text = kesilenmatadet.ToString("0.##");
                        }
                    }
                }
            }
            kesilentoplamm2 += kesilenacilmatm2 + kesilenparlakm2 + kesilenmatm2 + kesilenacilparlakm2;
            kesilentoplamadet += kesilenacilmatadet + kesilenmatadet + kesilenparlakadet + kesilenacilparlakadet;
            textBox2.Text = kesilentoplamm2.ToString();
            textBox1.Text = kesilentoplamadet.ToString();

        }
        string kesildimi_;
        DateTime tarih_;
        string renk_;
        string sipariştipi_;
        private void BugünKesilenlerrGüncel()
        {
            kesilenmatm2_ = 0;
            kesilenmatadet_ = 0;
            kesilenparlakm2_ = 0;
            kesilenparlakadet_ = 0;
            kesilenacilmatm2_ = 0;
            kesilenacilmatadet_ = 0;
            kesilenacilparlakm2_ = 0;
            kesilenacilparlakadet_ = 0;
            kesilentoplamm2_ = 0;
            kesilentoplamadet_ = 0;
            kesilentoplamm2bölüadet_ = 0;

            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT *FROM Siparişler where Onay=@Onay AND AnaSiparişMi=@p1 AND KesildiMi=@KesildiMi";
            komut.Parameters.AddWithValue("@Onay", "Onaylandı");
            komut.Parameters.AddWithValue("@p1", "Evet");
            komut.Parameters.AddWithValue("@KesildiMi", "Evet");
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                kesildimi_ = dr["KesildiMi"].ToString();
                renk_ = dr["Renk"].ToString();
                sipariştipi_ = dr["SiparişTipi"].ToString();

                //acil olmayan kısımlar

                if (dr["KesildiTarihi"].ToString() != "")
                {
                    tarih_ = Convert.ToDateTime(dr["KesildiTarihi"].ToString());
                    if (kesildimi_ == "Evet" && sipariştipi_ == "Acil" && tarih_.ToString("yyyy - MM - dd") == (DateTime.Now.ToString("yyyy - MM - dd")))
                    {
                        if (renk_.StartsWith("HG"))
                        {
                            kesilenacilparlakm2_ += Convert.ToDouble(dr["M2"]);
                            kesilenacilparlakadet_ += Convert.ToDouble(dr["Adet"]);
                            textBox31.Text = kesilenacilparlakm2_.ToString("0.##");
                            textBox32.Text = kesilenacilparlakadet_.ToString("0.##");
                        }
                        else
                        {
                            kesilenacilmatm2_ += Convert.ToDouble(dr["M2"]);
                            kesilenacilmatadet_ += Convert.ToDouble(dr["Adet"]);
                            textBox29.Text = kesilenacilmatm2_.ToString("0.##");
                            textBox30.Text = kesilenacilmatadet_.ToString("0.##");
                        }
                    }
                    else if (tarih_.ToString("yyyy - MM - dd") == (DateTime.Now.ToString("yyyy - MM - dd")) && kesildimi_ == "Evet")
                    {
                        if (renk_.StartsWith("HG"))
                        {
                            kesilenparlakm2_ += Convert.ToDouble(dr["M2"]);
                            kesilenparlakadet_ += Convert.ToDouble(dr["Adet"]);
                            textBox27.Text = kesilenparlakm2_.ToString("0.##");
                            textBox28.Text = kesilenparlakadet_.ToString("0.##");
                        }
                        else
                        {
                            kesilenmatm2_ += Convert.ToDouble(dr["M2"]);
                            kesilenmatadet_ += Convert.ToDouble(dr["Adet"]);
                            textBox17.Text = kesilenmatm2_.ToString("0.##");
                            textBox18.Text = kesilenmatadet_.ToString("0.##");
                        }
                    }
                }
            }
            kesilentoplamm2_ += kesilenacilmatm2_ + kesilenparlakm2_ + kesilenmatm2_ + kesilenacilparlakm2_;
            kesilentoplamadet_ += kesilenacilmatadet_ + kesilenmatadet_ + kesilenparlakadet_ + kesilenacilparlakadet_;
            kesilentoplamm2bölüadet_ = kesilentoplamm2_ / kesilentoplamadet_;
            textBox33.Text = kesilentoplamm2_.ToString();
            textBox34.Text = kesilentoplamadet_.ToString();
            textBox35.Text = kesilentoplamm2bölüadet_.ToString("0.##");

        }
        string acilsipariş;
        string kesilditarihi;
        private void AcilSipariş()
        {
            kesilditarihi = "";
            acilsipariş = "";
            if (dataGridView1.Rows.Count == 1)
            {
                timer2.Stop();
                label9.Text = "KESİLECEK SİPARİŞLER";
                label9.BackColor = Color.Blue;
            }
            else
            {
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Siparişler where Onay=@Onay AND AnaSiparişMi=@p1 AND KesildiTarihi is null";
                komut.Parameters.AddWithValue("@Onay", "Onaylandı");
                komut.Parameters.AddWithValue("@p1", "Evet");
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    string sipnu = dr["SiparisNo"].ToString();
                    acilsipariş = dr["SiparişTipi"].ToString();
                    kesilditarihi = dr["KesildiTarihi"].ToString();
                    if (acilsipariş == "Acil" && kesilditarihi.Length < 2)
                    {
                        timer2.Start();
                        label9.Text = "Dikkat! Acil Sipariş Var! Dikkat! Acil Sipariş Var! Dikkat! Acil Sipariş Var!";
                        break;
                    }
                    else
                    {
                        timer2.Stop();
                        label9.Text = "KESİLECEK SİPARİŞLER";
                        label9.TextAlign = ContentAlignment.MiddleCenter;
                    }
                }
            }
        }
        string siparisno,renkacil;
        private void TeslimTarihine3GünKalanlarıYakSöndür()
        {
            timer3.Stop();
            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT * FROM Siparişler WHERE CONVERT(datetime, TeslimTarihi, 104) <= DATEADD(day, 6, GETDATE()) AND Onay=@Onay AND AnaSiparişMi=@p1 AND KesildiTarihi is null ORDER BY SiparisNo ASC";
            komut.Parameters.AddWithValue("@Onay", "Onaylandı");
            komut.Parameters.AddWithValue("@p1", "Evet");
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                siparisno = dr["SiparisNo"].ToString();
                renkacil = dr["Renk"].ToString();
                timer3.Start();
                label9.Text += " / " + siparisno + "-" + renkacil + " / ";
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
        private void pictureBox8_Click(object sender, EventArgs e)
        {
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            label2.Text = DateTime.Now.ToLongDateString();
            label12.Text = DateTime.Now.ToLongTimeString();
        }

        private void methodlarıgetir()
        {
            liste();
            liste2();
            siparisnogetir();
            renkgetir();
            aynısipnugetirme();
            KesilmişKesilecekM2();
            KesilmişKesilecekAdet();
            MatParlakM2();
            MatParlakAdet();
            AcilMatParlakM2();
            AcilMatParlakAdet();
            KesilecekMatParlakM2();
            KesilecekMatParlakAdet();
            BugünKesilenlerr();
            BugünKesilenlerrGüncel();
            AcilSipariş();
            TeslimTarihine3GünKalanlarıYakSöndür();
            onayrenklendir();
        }
        private void Form6_Load(object sender, EventArgs e)
        {
            if (yetki == "CNC")
            {
                button21.Visible = false;
            }
            timer2.Enabled = true;
            timer1.Start();
            liste();
            methodlarıgetir();
            label18.Text += "\r\n"+ DateTime.Now.ToShortDateString();
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            label9.Text = label9.Text.Substring(1) + label9.Text.Substring(0, 1);
        }

        string izin;
        string izinsil;
        string hepsi;
        string paletlenen;
        string kesilecek;
        private void siparişkes()
        {
            izin = "";
            hepsi = "";
            paletlenen = "";
            kesilecek = "";
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString() == comboBox3.Text && dataGridView1.Rows[i].Cells[9].Value.ToString() == "")
                {
                    izin = "var";
                }
            }

            if (izin == "var")
            {
                string tarih = Convert.ToDateTime(DateTime.Now).ToString("yyyy-MM-dd HH:mm:ss");
                string sorgu = "UPDATE Siparişler SET KesildiTarihi=@KesildiTarihi,KesildiMi=@KesildiMi,Aşama=@Aşama WHERE SiparisNo=@SiparisNo";
                SqlCommand komut;
                komut = new SqlCommand(sorgu, bgl.baglanti());
                komut.Parameters.AddWithValue("@SiparisNo", comboBox3.Text);
                komut.Parameters.AddWithValue("@KesildiTarihi", tarih);
                komut.Parameters.AddWithValue("@KesildiMi", "Evet");
                komut.Parameters.AddWithValue("@Aşama", "Etiket");
                komut.ExecuteNonQuery();
                MessageBox.Show(comboBox3.Text + " 'lı sipariş kesilmiştir.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);


                SqlCommand komut2 = new SqlCommand();
                komut2.CommandText = "SELECT *From Grafik where Renk=@Renk";
                komut2.Parameters.AddWithValue("@Renk", dataGridView1.Rows[0].Cells["Renk"].Value.ToString());
                komut2.Connection = bgl.baglanti();
                komut2.CommandType = CommandType.Text;
                SqlDataReader dr;
                dr = komut2.ExecuteReader();
                while (dr.Read())
                {
                    hepsi = dr["Hepsi"].ToString();
                    paletlenen = dr["Paletlenen"].ToString();
                    kesilecek = dr["Paletlenecek"].ToString();
                }

                //string sorgu4 = "UPDATE Grafik SET Hepsi=@Hepsi, Paletlenecek=@Paletlenecek, Paletlenen=@Paletlenen WHERE Renk=@Renk";
                //SqlCommand komut4;
                //komut4 = new SqlCommand(sorgu4, bgl.baglanti());
                //komut4.Parameters.AddWithValue("@Renk", dataGridView1.Rows[0].Cells["Renk"].Value.ToString());
                //komut4.Parameters.AddWithValue("@Hepsi", Convert.ToString(Convert.ToDouble(hepsi) + Convert.ToDouble(dataGridView1.Rows[0].Cells["ToplamM2"].Value.ToString())));
                //komut4.Parameters.AddWithValue("@Paletlenecek", Convert.ToString(Convert.ToDouble(kesilecek) + Convert.ToDouble(dataGridView1.Rows[0].Cells["ToplamM2"].Value.ToString())));
                //komut4.Parameters.AddWithValue("@Paletlenen", Convert.ToString(Convert.ToDouble(paletlenen) + 0));
                //komut4.ExecuteNonQuery();
                label14.Text = comboBox3.Text;
                label17.Text = dataGridView1.Rows[0].Cells[3].Value.ToString();
                methodlarıgetir();
                
            }
            else
            {
                MessageBox.Show("Lütfen geçerli bir sipariş numarası giriniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void kesilensiparişsil()
        {
            izinsil = "";
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString() == comboBox3.Text && dataGridView1.Rows[i].Cells[9].Value.ToString() != "")
                {
                    izinsil = "var";
                }
            }

            if (izinsil == "var")
            {
                string sorgu = "UPDATE Siparişler SET KesildiTarihi=@KesildiTarihi,KesildiMi=@KesildiMi WHERE SiparisNo=@SiparisNo";
                SqlCommand komut;
                komut = new SqlCommand(sorgu, bgl.baglanti());
                komut.Parameters.AddWithValue("@SiparisNo", comboBox3.Text);
                komut.Parameters.Add("@KesildiTarihi", SqlDbType.DateTime).Value = DBNull.Value; // DateTime türünde bir alanı NULL olarak ayarla
                komut.Parameters.AddWithValue("@KesildiMi", "Hayır");
                komut.ExecuteNonQuery();
                MessageBox.Show(comboBox3.Text + " 'lu kesilen sipariş iptal edilmiştir.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                methodlarıgetir();
            }
            else
            {
                MessageBox.Show("Lütfen geçerli bir sipariş numarası giriniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void matsiparişler()
        {
            string srg = "HG%";
            string sorgu = "SELECT DISTINCT SiparisNo,TeslimTarihi,SiparişTarihi,Müşteri as'Firma Ünvanı',Model,Renk,SiparişTipi,ToplamM2,Onay,KesildiTarihi,ToplamAdet From Siparişler where Onay='Onaylandı' AND AnaSiparişMi='Evet' AND KesildiMi='Hayır' AND Renk NOT Like '" + srg + "'  ORDER BY SiparisNo DESC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            onayrenklendir();
        }
        private void parlaksiparişler()
        {
            string srg = "HG%";
            string sorgu = "SELECT DISTINCT SiparisNo,TeslimTarihi,SiparişTarihi,Müşteri as'Firma Ünvanı',Model,Renk,SiparişTipi,ToplamM2,Onay,KesildiTarihi,ToplamAdet From Siparişler where Onay='Onaylandı' AND AnaSiparişMi='Evet' AND KesildiMi='Hayır' AND Renk Like '" + srg + "'  ORDER BY SiparisNo DESC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            onayrenklendir();
        }
        private void bugünegöresırala()
        {
            DateTime bitir = DateTime.Now;
            DateTime basla = DateTime.Now;
            dateTimePicker1.Value = basla;
            label5.Text = basla.ToString("yyyy - MM - dd");
            label6.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");

            string sorgu = "SELECT DISTINCT SiparisNo,TeslimTarihi,SiparişTarihi,Müşteri as'Firma Ünvanı',Model,Renk,SiparişTipi,ToplamM2,Onay,KesildiTarihi,ToplamAdet From Siparişler where SiparişTarihi between '" + label5.Text + "' AND '" + label6.Text + "' AND  Onay='Onaylandı' AND AnaSiparişMi='Evet' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            onayrenklendir();
        }
        private void haftayagöresırala()
        {
            DateTime bitir = DateTime.Now;
            DateTime basla = bitir.AddDays(-7);
            dateTimePicker1.Value = basla;
            label5.Text = basla.ToString("yyyy - MM - dd HH:mm:ss");
            label6.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");

            string sorgu = "SELECT DISTINCT SiparisNo,TeslimTarihi,SiparişTarihi,Müşteri as'Firma Ünvanı',Model,Renk,SiparişTipi,ToplamM2,Onay,KesildiTarihi,ToplamAdet From Siparişler where SiparişTarihi between '" + label5.Text + "' AND '" + label6.Text + "' AND  Onay='Onaylandı' AND AnaSiparişMi='Evet' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            onayrenklendir();
        }
        private void ayagöresırala()
        {
            DateTime bitir = DateTime.Now;
            DateTime basla = bitir.AddMonths(-1);
            dateTimePicker1.Value = basla;
            label5.Text = basla.ToString("yyyy - MM - dd HH:mm:ss");
            label6.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");

            string sorgu = "SELECT DISTINCT SiparisNo,TeslimTarihi,SiparişTarihi,Müşteri as'Firma Ünvanı',Model,Renk,SiparişTipi,ToplamM2,Onay,KesildiTarihi,ToplamAdet From Siparişler where SiparişTarihi between '" + label5.Text + "' AND '" + label6.Text + "' AND  Onay='Onaylandı' AND AnaSiparişMi='Evet' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            onayrenklendir();
        }
        private void yılagöresırala()
        {
            DateTime bitir = DateTime.Now;
            DateTime basla = bitir.AddYears(-1);
            dateTimePicker1.Value = basla;
            label5.Text = basla.ToString("yyyy - MM - dd HH:mm:ss");
            label6.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");

            string sorgu = "SELECT DISTINCT SiparisNo,TeslimTarihi,SiparişTarihi,Müşteri as'Firma Ünvanı',Model,Renk,SiparişTipi,ToplamM2,Onay,KesildiTarihi,ToplamAdet From Siparişler where SiparişTarihi between '" + label5.Text + "' AND '" + label6.Text + "' AND  Onay='Onaylandı' AND AnaSiparişMi='Evet' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            onayrenklendir();
        }
        private void button6_Click(object sender, EventArgs e)
        {
            siparişkes();
            siparisnogetir();
        }

        private void comboBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                siparişkes();
                comboBox3.Text = "";
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            kesilensiparişsil();
            //methodlarıgetir();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            siparisnoyagöresırala();
            aynısipnugetirme();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            rengegöresırala();
            aynısipnugetirme();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            liste();
            aynısipnugetirme();
        }

        private void comboBox2_Click(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = -1;
            liste();
            aynısipnugetirme();
        }

        private void comboBox1_Click(object sender, EventArgs e)
        {
            comboBox2.SelectedIndex = -1;
            liste();
            aynısipnugetirme();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string sorgu = "SELECT DISTINCT SiparisNo, TeslimTarihi, SiparişTarihi, Müşteri as 'Firma Ünvanı', Model, Renk, SiparişTipi, ToplamM2, Onay, KesildiTarihi, ToplamAdet FROM Siparişler WHERE Onay='Onaylandı' AND AnaSiparişMi='Evet' AND KesildiMi='Hayır' ORDER BY TeslimTarihi ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            aynısipnugetirme();
            onayrenklendir();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form6 frm = new Form6();
            this.Hide();
            frm.yetki = yetki;
            frm.Show();
        }

        private void comboBox3_TextChanged(object sender, EventArgs e)
        {
            if (comboBox3.Text.Length == 0)
            {
                liste();
                aynısipnugetirme();
            }
            else
            {
                string srg = comboBox3.Text;
                string sorgu = "Select DISTINCT SiparisNo,TeslimTarihi,SiparişTarihi,Müşteri as'Firma Ünvanı',Model,Renk,SiparişTipi,ToplamM2,Onay,KesildiTarihi,ToplamAdet from Siparişler where SiparisNo Like '" + srg + "' AND AnaSiparişMi='" + "Evet" + "' AND Onay='" + "Onaylandı" + "'";
                SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
                DataSet ds = new DataSet();
                adap.Fill(ds, "Siparişler");
                this.dataGridView1.DataSource = ds.Tables[0];
                aynısipnugetirme();
                try
                {
                    SqlCommand komut = new SqlCommand();
                    komut.CommandText = "SELECT * FROM Paletler where Renk=@Renk";
                    komut.Parameters.AddWithValue("@Renk", dataGridView1.Rows[0].Cells["Renk"].Value);
                    komut.Connection = bgl.baglanti();
                    komut.CommandType = CommandType.Text;

                    SqlDataReader dr;
                    dr = komut.ExecuteReader();
                    while (dr.Read())
                    {
                        textBox16.Text = dr["Palet"].ToString();
                    }
                }
                catch (Exception)
                {
                    textBox16.Text =" ";
                }

            }




        }

        private void button3_Click(object sender, EventArgs e)
        {
            matsiparişler();
            aynısipnugetirme();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            parlaksiparişler();
            aynısipnugetirme();
        }

        private void button18_Click(object sender, EventArgs e)
        {
            bugünegöresırala();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            haftayagöresırala();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            ayagöresırala();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            yılagöresırala();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTime bitir = dateTimePicker1.Value;
            DateTime basla = dateTimePicker2.Value;
            label5.Text = basla.ToString("yyyy - MM - dd");
            label6.Text = bitir.ToString("yyyy - MM - dd");

            string sorgu = "SELECT DISTINCT SiparisNo,TeslimTarihi,SiparişTarihi,Müşteri as'Firma Ünvanı',Model,Renk,SiparişTipi,ToplamM2,Onay,KesildiTarihi,ToplamAdet From Siparişler where SiparişTarihi between '" + label6.Text + "' AND '" + label5.Text + "' AND  Onay='Onaylandı' AND AnaSiparişMi='Evet' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            onayrenklendir();
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            DateTime bitir = dateTimePicker1.Value;
            DateTime basla = dateTimePicker2.Value;
            label5.Text = basla.ToString("yyyy - MM - dd");
            label6.Text = bitir.ToString("yyyy - MM - dd");

            string sorgu = "SELECT DISTINCT SiparisNo,TeslimTarihi,SiparişTarihi,Müşteri as'Firma Ünvanı',Model,Renk,SiparişTipi,ToplamM2,Onay,KesildiTarihi,ToplamAdet From Siparişler where SiparişTarihi between '" + label6.Text + "' AND '" + label5.Text + "' AND  Onay='Onaylandı' AND AnaSiparişMi='Evet' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            onayrenklendir();
        }

        private void Form6_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Hide();
            Form1 form1 = Application.OpenForms["Form1"] as Form1;
            if (form1 != null)
            {
                form1.Show();
            }
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void dataGridView1_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            aynısipnugetirme();
        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void timer3_Tick(object sender, EventArgs e)
        {
        }
        string sipn;
        int satir2;
        int sipnosayısı;
        int x;
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
            sheet2.Cells[4, 4].Value = dataGridView2.Rows[satir2].Cells["Firma Ünvanı"].Value.ToString(); // müşteri yazdırma
            sheet2.Cells[5, 4].Value = dataGridView2.Rows[satir2].Cells["Adres"].Value.ToString(); // adres yazdırma
            sheet2.Cells[7, 4].Value = dataGridView2.Rows[satir2].Cells["Telefon"].Value.ToString(); // telefon yazdırma
            sheet2.Cells[7, 6].Value = dataGridView2.Rows[satir2].Cells["Telefon"].Value.ToString(); // telefon yazdırma
            sheet2.Cells[7, 9].Value = dataGridView2.Rows[satir2].Cells["SiparişTarihi"].Value.ToString(); // sip tarih yazdırma
            sheet2.Cells[8, 9].Value = dataGridView2.Rows[satir2].Cells["TeslimTarihi"].Value.ToString(); // tes tarih yazdırma


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

        private void button23_Click(object sender, EventArgs e)
        {
            Form18 frm = new Form18();
            frm.yetki = yetki;
            frm.ShowDialog();
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
                frm.hangiformdan = "Form6";
                frm.yetki = yetki;
                this.Hide();
                frm.ShowDialog();
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            Form2 frm = new Form2();
            frm.yetki = yetki;
            this.Hide();
            frm.Show();
        }
    }
}
