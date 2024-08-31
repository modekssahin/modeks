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
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using System.IO;
using System.Diagnostics;
using Newtonsoft.Json.Linq;

namespace Modeks
{
    public partial class Form7 : Form
    {
        public Form7()
        {
            InitializeComponent();
        }
        sqlsinif bgl = new sqlsinif();
        public string yetki;
        double matm2 = 0;
        double parlakm2 = 0;
        double matadet = 0;
        double parlakadet = 0;
        double acilmatm2 = 0;
        double acilparlakm2 = 0;
        double acilmatadet = 0;
        double acilparlakadet = 0;
        double paletyapılacakmatparlakm2 = 0;
        double paletyapılacakmatparlakadet = 0;

        double paletolanmatm2 = 0;
        double paletolanparlakm2 = 0;
        double paletolanmatadet = 0;
        double paletolanparlakadet = 0;
        double paletolanacilmatm2 = 0;
        double paletolanacilparlakm2 = 0;
        double paletolanacilmatadet = 0;
        double paletolanacilparlakadet = 0;
        double paletolanmatparlakm2 = 0;
        double paletolankmatparlakadet = 0;
        double paletolantoplamm2 = 0;
        double paletolantoplamadet = 0;

        double tümsiparişm2 = 0;
        double tümsiparişadet = 0;
        double paletolansayısı = 0;
        double paletolanm2 = 0;
        double paletolanadet = 0;
        double paletolacaksayısı = 0;

        double kesileceksayısı = 0;
        double kesilecekm2 = 0;
        double kesilecekadet = 0;
        double kesildipaletyapsayısı = 0;
        double kesildipaletyapm2 = 0;
        double kesildipaletyapadet = 0;
        double onaybekleyensayısı = 0;
        double onaybekleyenm2 = 0;
        double onaybekleyenadet = 0;


        double kayitsayisi = 0;
        public void liste()
        {
            string kayit = "SELECT DISTINCT SiparisNo,Müşteri,Model,Renk,SiparişTipi,SiparişTarihi,KesildiMi,KesildiTarihi,Onay,Palet,Etiket,ToplamM2,ToplamAdet,TeslimTarihi From Siparişler where  AnaSiparişMi=@p1 AND Aşama=@Aşama ORDER BY SiparisNo DESC";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            komut.Parameters.AddWithValue("@p1", "Evet");
            komut.Parameters.AddWithValue("@Aşama", "Etiket");
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            üçgüniçindeki();
            bgl.baglanti().Close();
            onayrenklendir();
        }
        public void liste2()
        {
            string kayit2 = "SELECT SiparisNo,Müşteri,Model,Renk,SiparişTipi,SiparişTarihi,KesildiMi,KesildiTarihi,Onay,Palet,Etiket,ToplamM2,ToplamAdet,Boy,En,Özellik,Telefon,TeslimTarihi,Adres,M2,Adet From Siparişler where Aşama=@Aşama ORDER BY SiparisNo ASC,Boy DESC";
            SqlCommand komut2 = new SqlCommand(kayit2, bgl.baglanti());
            komut2.Parameters.AddWithValue("@Aşama", "Etiket");
            SqlDataAdapter da2 = new SqlDataAdapter(komut2);
            DataTable dt2 = new DataTable();
            da2.Fill(dt2);
            dataGridView2.DataSource = dt2;
            //dataGridView2.Columns["Boy"].Visible = false;
            //dataGridView2.Columns["En"].Visible = false;
            //dataGridView2.Columns["Özellik"].Visible = false;
            //dataGridView2.Columns["Telefon"].Visible = false;
            //dataGridView2.Columns["TeslimTarihi"].Visible = false;
            //dataGridView2.Columns["Adres"].Visible = false;
            bgl.baglanti().Close();
            onayrenklendir();
        }
        private void MatParlakM2()
        {
            try
            {
                matm2 = 0;
                parlakm2 = 0;
                for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
                {
                    if (dataGridView1.Rows[i].Visible == true)
                        if (dataGridView1.Rows[i].Cells[3].Value.ToString().StartsWith("HG") && (dataGridView1.Rows[i].Cells[4].Value.ToString() == "Normal" || dataGridView1.Rows[i].Cells[4].Value.ToString() == "Üretim Sorunu"))
                            if (dataGridView1.Rows[i].Cells[9].Value.ToString().Length < 1)
                                parlakm2 += Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value);
                            else
                            {

                            }
                        else if (dataGridView1.Rows[i].Cells[9].Value.ToString().Length < 1 && (dataGridView1.Rows[i].Cells[4].Value.ToString() == "Normal" || dataGridView1.Rows[i].Cells[4].Value.ToString() == "Üretim Sorunu"))
                            matm2 += Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value);
                }
                textBox5.Text = matm2.ToString("0.##");
                textBox8.Text = parlakm2.ToString("0.##");
            }
            catch (Exception ex)
            {
                MessageBox.Show("MatParlakM2 : " + ex.Message);
            }

        }
        private void MatParlakAdet()
        {
            try
            {
                matadet = 0;
                parlakadet = 0;
                for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
                {
                    if (dataGridView1.Rows[i].Visible == true)
                        if (dataGridView1.Rows[i].Cells[3].Value.ToString().StartsWith("HG") && (dataGridView1.Rows[i].Cells[4].Value.ToString() == "Normal" || dataGridView1.Rows[i].Cells[4].Value.ToString() == "Üretim Sorunu"))
                            if (dataGridView1.Rows[i].Cells[9].Value.ToString().Length < 1)
                                parlakadet += Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value);
                            else
                            {

                            }
                        else if (dataGridView1.Rows[i].Cells[9].Value.ToString().Length < 1 && (dataGridView1.Rows[i].Cells[4].Value.ToString() == "Normal" || dataGridView1.Rows[i].Cells[4].Value.ToString() == "Üretim Sorunu"))

                            matadet += Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value);
                }
                textBox6.Text = matadet.ToString("0.##");
                textBox7.Text = parlakadet.ToString("0.##");
            }
            catch (Exception ex)
            {
                MessageBox.Show("MatParlakAdet : " + ex.Message);
            }

        }
        private void AcilMatParlakM2()
        {
            try
            {
                acilmatm2 = 0;
                acilparlakm2 = 0;
                for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
                {
                    if (dataGridView1.Rows[i].Visible == true)
                        if (dataGridView1.Rows[i].Cells[3].Value.ToString().StartsWith("HG") && dataGridView1.Rows[i].Cells[4].Value.ToString() == "Acil")
                            if (dataGridView1.Rows[i].Cells[9].Value.ToString().Length < 1)
                                acilparlakm2 += Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value);
                            else
                            {

                            }
                        else if (dataGridView1.Rows[i].Cells[4].Value.ToString() == "Acil" && dataGridView1.Rows[i].Cells[9].Value.ToString().Length < 1)
                            acilmatm2 += Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value);
                }
                textBox10.Text = acilmatm2.ToString("0.##");
                textBox12.Text = acilparlakm2.ToString("0.##");
            }
            catch (Exception ex)
            {
                MessageBox.Show("AcilMatParlakM2 : " + ex.Message);
            }

        }
        private void AcilMatParlakAdet()
        {
            try
            {
                acilmatadet = 0;
                acilparlakadet = 0;
                for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
                {
                    if (dataGridView1.Rows[i].Visible == true)
                        if (dataGridView1.Rows[i].Cells[3].Value.ToString().StartsWith("HG") && dataGridView1.Rows[i].Cells[4].Value.ToString() == "Acil")
                            if (dataGridView1.Rows[i].Cells[9].Value.ToString().Length < 1)
                                acilparlakadet += Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value);
                            else
                            {

                            }
                        else if (dataGridView1.Rows[i].Cells[4].Value.ToString() == "Acil" && dataGridView1.Rows[i].Cells[9].Value.ToString().Length < 1)
                            acilmatadet += Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value);
                }
                textBox9.Text = acilmatadet.ToString("0.##");
                textBox11.Text = acilparlakadet.ToString("0.##");
            }
            catch (Exception ex)
            {
                MessageBox.Show("AcilMatParlakAdet : " + ex.Message);
            }

        }
        private void PaletYapılacakMatParlakM2()
        {
            paletyapılacakmatparlakm2 = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true)
                    if (dataGridView1.Rows[i].Cells[9].Value.ToString().Length < 1)
                        paletyapılacakmatparlakm2 += Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value);
            }
            textBox14.Text = paletyapılacakmatparlakm2.ToString("0.##");
            textBox35.Text = paletyapılacakmatparlakm2.ToString("0.##");
        }
        private void PaletYapılacakMatParlakAdet()
        {
            paletyapılacakmatparlakadet = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true)
                    if (dataGridView1.Rows[i].Cells[9].Value.ToString().Length < 1)
                        paletyapılacakmatparlakadet += Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value);
            }
            textBox13.Text = paletyapılacakmatparlakadet.ToString("0.##");
            textBox43.Text = paletyapılacakmatparlakadet.ToString("0.##");
        }

        private void TümSiparişM2()
        {
            tümsiparişm2 = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true)
                    tümsiparişm2 += Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value);
            }
            textBox40.Text = tümsiparişm2.ToString("0.##");
            textBox39.Text = tümsiparişm2.ToString("0.##");
        }
        private void TümSiparişAdet()
        {
            tümsiparişadet = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true)
                    tümsiparişadet += Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value);
            }
            textBox48.Text = tümsiparişadet.ToString("0.##");
            textBox47.Text = tümsiparişadet.ToString("0.##");
        }
        private void PaletOlanSayısı()
        {
            paletolansayısı = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true && dataGridView1.Rows[i].Cells[9].Value.ToString().Length >= 1)
                    paletolansayısı++;
            }
            textBox28.Text = paletolansayısı.ToString("0.##");
        }
        private void PaletOlanM2()
        {
            paletolanm2 = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true && dataGridView1.Rows[i].Cells[9].Value.ToString().Length >= 1)
                    paletolanm2 += Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value);
            }
            textBox37.Text = paletolanm2.ToString("0.##");
        }
        private void PaletOlanAdet()
        {
            paletolanadet = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true && dataGridView1.Rows[i].Cells[9].Value.ToString().Length >= 1)
                    paletolanadet += Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value);
            }
            textBox45.Text = paletolanadet.ToString("0.##");
        }
        private void PaletOlacakSayısı()
        {
            paletolacaksayısı = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true && dataGridView1.Rows[i].Cells[9].Value.ToString().Length < 1)
                    paletolacaksayısı++;
            }
            textBox30.Text = paletolacaksayısı.ToString("0.##");
        }
        double kesilenm2 = 0;
        double kesilenadet = 0;
        private void KesilenM2()
        {
            kesilenm2 = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true && dataGridView1.Rows[i].Cells[6].Value.ToString() == "Evet")
                    kesilenm2 += Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value);
            }
            textBox39.Text = kesilenm2.ToString("0.##");
        }
        private void KesilenAdet()
        {
            kesilenadet = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true && dataGridView1.Rows[i].Cells[6].Value.ToString() == "Evet")
                    kesilenadet += Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value);
            }
            textBox47.Text = kesilenadet.ToString("0.##");
        }
        private void KesilecekM2()
        {
            kesilecekm2 = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true && dataGridView1.Rows[i].Cells[6].Value.ToString() == "Hayır")
                    kesilecekm2 += Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value);
            }
            textBox38.Text = kesilecekm2.ToString("0.##");
        }
        private void KesilecekAdet()
        {
            kesilecekadet = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true && dataGridView1.Rows[i].Cells[6].Value.ToString() == "Hayır")
                    kesilecekadet += Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value);
            }
            textBox46.Text = kesilecekadet.ToString("0.##");
        }
        int siparişsayısı = 0;
        private void SiparişSayısı()
        {
            siparişsayısı = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true)
                    siparişsayısı++;
            }
            textBox17.Text = siparişsayısı.ToString("0.##");
        }
        int kesilensayısı = 0;
        private void KesilenSayısı()
        {
            kesilensayısı = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true && dataGridView1.Rows[i].Cells[6].Value.ToString() == "Evet")
                    kesilensayısı++;
            }
            textBox18.Text = kesilensayısı.ToString("0.##");
        }
        private void KesilecekSayısı()
        {
            kesileceksayısı = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true && dataGridView1.Rows[i].Cells[6].Value.ToString() == "Hayır")
                    kesileceksayısı++;
            }
            textBox27.Text = kesileceksayısı.ToString("0.##");
        }
        private void KesildiPaletYapM2()
        {
            kesildipaletyapm2 = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true && dataGridView1.Rows[i].Cells[6].Value.ToString() == "Evet" && dataGridView1.Rows[i].Cells[9].Value.ToString().Length < 1)
                    kesildipaletyapm2 += Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value);
            }
            textBox36.Text = kesildipaletyapm2.ToString("0.##");
        }
        private void KesildiPaletYapAdet()
        {
            kesildipaletyapadet = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true && dataGridView1.Rows[i].Cells[6].Value.ToString() == "Evet" && dataGridView1.Rows[i].Cells[9].Value.ToString().Length < 1)
                    kesildipaletyapadet += Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value);
            }
            textBox44.Text = kesildipaletyapadet.ToString("0.##");
        }
        private void KesildiPaletYapSayısı()
        {
            kesildipaletyapsayısı = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true && dataGridView1.Rows[i].Cells[6].Value.ToString() == "Evet" && dataGridView1.Rows[i].Cells[9].Value.ToString().Length < 1)
                    kesildipaletyapsayısı++;
            }
            textBox29.Text = kesildipaletyapsayısı.ToString("0.##");
        }
        private void OnayBekleyenM2()
        {
            onaybekleyenm2 = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true && (dataGridView1.Rows[i].Cells[8].Value.ToString() == "Onay Bekliyor" || dataGridView1.Rows[i].Cells[8].Value.ToString() == "Onay Bekliyor "))
                    onaybekleyenm2 += Convert.ToDouble(dataGridView1.Rows[i].Cells[11].Value);
            }
            textBox34.Text = onaybekleyenm2.ToString("0.##");
        }
        private void OnayBekleyenAdet()
        {
            onaybekleyenadet = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true && (dataGridView1.Rows[i].Cells[8].Value.ToString() == "Onay Bekliyor" || dataGridView1.Rows[i].Cells[8].Value.ToString() == "Onay Bekliyor "))
                    onaybekleyenadet += Convert.ToDouble(dataGridView1.Rows[i].Cells[12].Value);
            }
            textBox42.Text = onaybekleyenadet.ToString("0.##");
        }
        private void OnayBekleyenSayısı()
        {
            onaybekleyensayısı = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true && (dataGridView1.Rows[i].Cells[8].Value.ToString() == "Onay Bekliyor" || dataGridView1.Rows[i].Cells[8].Value.ToString() == "Onay Bekliyor "))
                    onaybekleyensayısı++;
            }
            textBox31.Text = onaybekleyensayısı.ToString("0.##");
        }
        private void kayıtsayısı()
        {
            try
            {
                SqlCommand cmd = new SqlCommand("select COUNT(DISTINCT SiparisNo) from Siparişler where AnaSiparişMi='" + "Evet" + "' AND Aşama='" + "Etiket" + "'", bgl.baglanti());
                kayitsayisi = Convert.ToInt32(cmd.ExecuteScalar());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kayit Sayisi : " + ex.Message);
            }
        }

        private void üçgüniçindeki()
        {
            try
            {
                label24.Text = DateTime.Now.Day.ToString();
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    if ((dataGridView1.Rows[i].Cells[8].Value.ToString() == "Onay Bekliyor" || dataGridView1.Rows[i].Cells[8].Value.ToString() == "Onay Bekliyor "))
                    {
                        //string kelime = "03.09.2023 21:28:50";
                        string kelime = dataGridView1.Rows[i].Cells["SiparişTarihi"].Value.ToString();

                        string[] parcalanmisKelime = kelime.Split('.');

                        string tarih = parcalanmisKelime[0].PadLeft(2, '0');
                        label23.Text = tarih.ToString();
                        if (Convert.ToInt32(label24.Text) > Convert.ToInt32(label23.Text))
                        {
                            if (Convert.ToInt32(label24.Text) - Convert.ToInt32(label23.Text) > 10)
                            {
                                dataGridView1.CurrentCell = null;
                                dataGridView1.Rows[i].Visible = false;
                            }
                        }
                        else if (Convert.ToInt32(label23.Text) - Convert.ToInt32(label24.Text) > 10)
                        {
                            dataGridView1.CurrentCell = null;
                            dataGridView1.Rows[i].Visible = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Üç Gün İçindeki : " + ex.Message);
            }

        }
        string renk;
        DateTime tarih;
        string sipariştipi;
        private void BugünBasılanlar()
        {
            paletolanmatm2 = 0;
            paletolanparlakm2 = 0;
            paletolanmatadet = 0;
            paletolanparlakadet = 0;
            paletolanacilmatm2 = 0;
            paletolanacilparlakm2 = 0;
            paletolanacilmatadet = 0;
            paletolanacilparlakadet = 0;

            paletolanmatparlakm2 = 0;
            paletolankmatparlakadet = 0;

            paletolantoplamm2 = 0;
            paletolantoplamadet = 0;
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
                if (dr["Etiket"].ToString() != "")
                {
                    tarih = Convert.ToDateTime(dr["Etiket"].ToString());
                    sipariştipi = dr["SiparişTipi"].ToString();

                    if (renk.StartsWith("HG") && tarih.ToString("yyyy - MM - dd") == (DateTime.Now.ToString("yyyy - MM - dd")))
                    {
                        if (sipariştipi == "Acil")
                        {
                            paletolanacilparlakm2 += Convert.ToDouble(dr["M2"]);
                            paletolanacilparlakadet += Convert.ToDouble(dr["Adet"]);
                            textBox20.Text = paletolanacilparlakm2.ToString("0.##");
                            textBox19.Text = paletolanacilparlakadet.ToString("0.##");
                        }
                        else if (sipariştipi == "Normal" || sipariştipi == "Üretim Sorunu")
                        {
                            paletolanparlakm2 += Convert.ToDouble(dr["M2"]);
                            paletolanparlakadet += Convert.ToDouble(dr["Adet"]);

                            textBox24.Text = paletolanparlakm2.ToString("0.##");
                            textBox23.Text = paletolanparlakadet.ToString("0.##");
                        }


                    }
                    else if (!renk.StartsWith("HG") && tarih.ToString("yyyy - MM - dd") == (DateTime.Now.ToString("yyyy - MM - dd")))
                    {
                        if (sipariştipi == "Acil")
                        {
                            paletolanacilmatm2 += Convert.ToDouble(dr["M2"]);
                            paletolanacilmatadet += Convert.ToDouble(dr["Adet"]);

                            textBox22.Text = paletolanacilmatm2.ToString("0.##");
                            textBox21.Text = paletolanacilmatadet.ToString("0.##");
                        }
                        else if (sipariştipi == "Normal" || sipariştipi == "Üretim Sorunu")
                        {
                            paletolanmatm2 += Convert.ToDouble(dr["M2"]);
                            paletolanmatadet += Convert.ToDouble(dr["Adet"]);

                            textBox26.Text = paletolanmatm2.ToString("0.##");
                            textBox25.Text = paletolanmatadet.ToString("0.##");
                        }
                    }
                }
            }
            paletolanmatparlakm2 += paletolanmatm2 + paletolanparlakm2;
            paletolankmatparlakadet += paletolanmatadet + paletolanparlakadet;
            textBox3.Text = paletolanmatparlakm2.ToString("0.##");
            textBox4.Text = paletolankmatparlakadet.ToString("0.##");

            paletolantoplamm2 += paletolanmatm2 + paletolanacilmatm2 + paletolanparlakm2 + paletolanacilparlakm2;
            paletolantoplamadet += paletolanmatadet + paletolanacilmatadet + paletolanparlakadet + paletolanacilparlakadet;
            textBox52.Text = paletolantoplamm2.ToString();
            textBox2.Text = paletolantoplamm2.ToString();
            textBox51.Text = paletolantoplamadet.ToString();
            textBox1.Text = paletolantoplamadet.ToString();
            textBox50.Text = (paletolantoplamadet / paletolantoplamm2).ToString();
            bgl.baglanti().Close();

        }
        string acilsipariş;
        string kesilditarihi;
        string Etiket;
        string Aşama;
        public void AcilSipariş()
        {
            kesilditarihi = "";
            acilsipariş = "";
            Aşama = "";

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
                string sipnu = dr["SiparisNo"].ToString();
                acilsipariş = dr["SiparişTipi"].ToString();
                kesilditarihi = dr["KesildiTarihi"].ToString();
                Etiket = dr["Etiket"].ToString();
                Aşama = dr["Aşama"].ToString();
                if ((acilsipariş == "Acil" && kesilditarihi.Length < 2) || (acilsipariş == "Acil" && Etiket.Length < 2) || (acilsipariş == "Acil" && Aşama == "Etiket"))
                {
                    timer2.Start();
                    label9.Text = "Dikkat! Acil Sipariş Var! Dikkat! Acil Sipariş Var! Dikkat! Acil Sipariş Var!";
                    label9.BackColor = Color.Red;
                    break;
                }
                else
                {
                    timer2.Stop();
                    label9.Text = "PALET OLACAK SİPARİŞLER";
                    label9.TextAlign = ContentAlignment.MiddleCenter;
                    label9.BackColor = Color.Blue;
                }
            }
            bgl.baglanti().Close();

        }
        private void siparisnogetir()
        {
            comboBox1.Items.Clear();
            comboBox3.Items.Clear();
            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT DISTINCT SiparisNo FROM Siparişler where AnaSiparişMi='" + "Evet" + "' AND Aşama='" + "Etiket" + "' ORDER BY SiparisNo ASC";
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                comboBox1.Items.Add(dr["SiparisNo"]);
                comboBox3.Items.Add(dr["SiparisNo"]);
            }
            bgl.baglanti().Close();

        }
        private void siparisnoyagöresırala()
        {
            string srg = comboBox1.Text;
            string sorgu = "SELECT DISTINCT SiparisNo,Müşteri,Model,Renk,SiparişTipi,SiparişTarihi,KesildiMi,KesildiTarihi,Onay,Palet,Etiket,M2,ToplamAdet,Boy,En,Özellik,Telefon,TeslimTarihi,Adres From Siparişler where SiparisNo Like '" + srg + "' AND AnaSiparişMi='" + "Evet" + "' AND Aşama='" + "Etiket" + "' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            dataGridView1.Columns["Boy"].Visible = false;
            dataGridView1.Columns["En"].Visible = false;
            dataGridView1.Columns["Özellik"].Visible = false;
            dataGridView1.Columns["Telefon"].Visible = false;
            dataGridView1.Columns["TeslimTarihi"].Visible = false;
            dataGridView1.Columns["Adres"].Visible = false;
            bgl.baglanti().Close();

        }
        private void renkgetir()
        {
            comboBox2.Items.Clear();
            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT DISTINCT Renk FROM Siparişler where AnaSiparişMi='" + "Evet" + "' AND Aşama='" + "Etiket" + "' ORDER BY Renk ASC";
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                comboBox2.Items.Add(dr["Renk"]);
            }
            bgl.baglanti().Close();

        }
        private void rengegöresırala()
        {
            string srg = comboBox2.Text;
            string sorgu = "SELECT DISTINCT SiparisNo,Müşteri,Model,Renk,SiparişTipi,SiparişTarihi,KesildiMi,KesildiTarihi,Onay,Palet,Etiket,M2,ToplamAdet,Boy,En,Özellik,Telefon,TeslimTarihi,Adres From Siparişler where Renk Like '" + srg + "' AND AnaSiparişMi='" + "Evet" + "' AND Aşama='" + "Etiket" + "' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            dataGridView1.Columns["Boy"].Visible = false;
            dataGridView1.Columns["En"].Visible = false;
            dataGridView1.Columns["Özellik"].Visible = false;
            dataGridView1.Columns["Telefon"].Visible = false;
            dataGridView1.Columns["TeslimTarihi"].Visible = false;
            dataGridView1.Columns["Adres"].Visible = false;
            bgl.baglanti().Close();

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
                onayrenklendir();
            }
            catch (Exception ex)
            {
                MessageBox.Show("AynıSipNu : " + ex.Message);
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

        string kesildiMi;
        private void Methodlar()
        {
            kesildiMi = "";
            kayıtsayısı();
            liste();
            liste2();

            üçgüniçindeki();
            aynısipnugetirme();
            GrafikDeğerleriGüncelle();
            MatParlakM2();
            MatParlakAdet();
            AcilMatParlakM2();
            AcilMatParlakAdet();
            PaletYapılacakMatParlakM2();
            PaletYapılacakMatParlakAdet();

            BugünBasılanlar();
            PressYollanan();

            TümSiparişM2();
            TümSiparişAdet();
            PaletOlanSayısı();
            PaletOlanM2();
            PaletOlanAdet();
            PaletOlacakSayısı();

            SiparişSayısı();
            KesilenSayısı();
            KesilecekM2();
            KesilecekAdet();
            KesilenM2();
            KesilenAdet();
            KesilecekSayısı();
            KesildiPaletYapM2();
            KesildiPaletYapAdet();
            KesildiPaletYapSayısı();
            OnayBekleyenM2();
            OnayBekleyenAdet();
            OnayBekleyenSayısı();

            AcilSipariş();
            TeslimTarihine3GünKalanlarıYakSöndür();
            siparisnogetir();
            renkgetir();
            onayrenklendir();
            bgl.baglanti().Close();


        }
        double grafik_2_paletlenen;
        double grafik_2_paletlenecek;
        double grafik_2_hepsi;
        string grafik_2_palet;
        string grafik_2_onay_bekleyen;
        private void GrafikDeğerleriGüncelle()
        {
            try
            {
                using (SqlConnection connection = bgl.baglanti())
                {
                    string updateQuery = "UPDATE Grafik SET Hepsi = @Hepsi, Paletlenecek = @Paletlenecek, Paletlenen = @Paletlenen";

                    using (SqlCommand updateCommand = new SqlCommand(updateQuery, connection))
                    {
                        updateCommand.Parameters.AddWithValue("@Hepsi", "0");
                        updateCommand.Parameters.AddWithValue("@Paletlenecek", "0");
                        updateCommand.Parameters.AddWithValue("@Paletlenen", "0");
                        updateCommand.ExecuteNonQuery();
                    }
                }

                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    using (SqlConnection connection = bgl.baglanti())
                    {
                        string selectQuery = "SELECT * FROM Grafik WHERE Renk = @Renk";

                        using (SqlCommand selectCommand = new SqlCommand(selectQuery, connection))
                        {
                            selectCommand.Parameters.AddWithValue("@Renk", dataGridView1.Rows[i].Cells["Renk"].Value.ToString());

                            using (SqlDataReader dr = selectCommand.ExecuteReader())
                            {
                                while (dr.Read())
                                {
                                    grafik_2_paletlenen = Convert.ToDouble(dr["Paletlenen"].ToString());
                                    grafik_2_paletlenecek = Convert.ToDouble(dr["Paletlenecek"].ToString());
                                    grafik_2_hepsi = Convert.ToDouble(dr["Hepsi"].ToString());
                                    grafik_2_palet = dr["Palet"].ToString();
                                }
                            }
                        }

                        if (dataGridView1.Rows[i].Visible)
                        {
                            string updateRowQuery = "UPDATE Grafik SET Hepsi = @Hepsi, Paletlenecek = @Paletlenecek, Paletlenen = @Paletlenen, Palet = @Palet WHERE Renk = @Renk";

                            using (SqlCommand updateRowCommand = new SqlCommand(updateRowQuery, connection))
                            {
                                updateRowCommand.Parameters.AddWithValue("@Renk", dataGridView1.Rows[i].Cells["Renk"].Value.ToString());

                                double hepsiValue = Convert.ToDouble(grafik_2_hepsi) + Convert.ToDouble(dataGridView1.Rows[i].Cells["ToplamM2"].Value.ToString());
                                updateRowCommand.Parameters.AddWithValue("@Hepsi", hepsiValue);

                                double paletlenenValue = dataGridView1.Rows[i].Cells["Palet"].Value.ToString().Length >= 1 && dataGridView1.Rows[i].Cells["Etiket"].Value.ToString().Length >= 3
                                    ? Convert.ToDouble(dataGridView1.Rows[i].Cells["ToplamM2"].Value.ToString())
                                    : 0;
                                updateRowCommand.Parameters.AddWithValue("@Paletlenen", grafik_2_paletlenen + paletlenenValue);

                                double paletlenecekValue = dataGridView1.Rows[i].Cells["Palet"].Value.ToString().Length < 1 || dataGridView1.Rows[i].Cells["Etiket"].Value.ToString().Length <= 3
                                    ? Convert.ToDouble(dataGridView1.Rows[i].Cells["ToplamM2"].Value.ToString())
                                    : 0;
                                updateRowCommand.Parameters.AddWithValue("@Paletlenecek", grafik_2_paletlenecek + paletlenecekValue);
                                if (grafik_2_palet == null)
                                    grafik_2_palet = "";
                                updateRowCommand.Parameters.AddWithValue("@Palet", grafik_2_palet.Replace(",", "."));


                                updateRowCommand.ExecuteNonQuery();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Grafik : " + ex.Message);
            }
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
        private void Form7_Load(object sender, EventArgs e)
        {
            radioButton1.Checked = true;
            if (yetki == "Etiket ve Palet")
            {
                button28.Visible = false;
            }
            timer1.Start();
            Methodlar();
            bgl.baglanti().Close();

        }

        private void pictureBox8_Click(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            label2.Text = DateTime.Now.ToLongDateString();
            label12.Text = DateTime.Now.ToLongTimeString();

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            siparisnoyagöresırala();
            aynısipnugetirme();
        }

        private void comboBox3_TextChanged(object sender, EventArgs e)
        {
            textBox16.Enabled = true;
            if (comboBox3.Text.Length == 0)
            {
                liste();
                liste2();
                aynısipnugetirme();
            }
            else
            {
                string srg = comboBox3.Text;
                string sorgu = "SELECT DISTINCT SiparisNo,Müşteri,Model,Renk,SiparişTipi,SiparişTarihi,KesildiMi,KesildiTarihi,Onay,Palet,Etiket,ToplamM2,ToplamAdet From Siparişler where SiparisNo Like '" + srg + "' AND  AnaSiparişMi='" + "Evet" + "' ORDER BY SiparisNo ASC";
                SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
                DataSet ds = new DataSet();
                adap.Fill(ds, "Siparişler");
                this.dataGridView1.DataSource = ds.Tables[0];
                aynısipnugetirme();

                if (dataGridView1.Rows.Count > 1)
                {
                    string srg3 = comboBox3.Text;
string sorgu3 = $@"
SELECT SiparisNo,
       Müşteri,
       Model,
       Renk,
       SiparişTipi,
       SiparişTarihi,
       KesildiMi,
       KesildiTarihi,
       Onay,
       Palet,
       Etiket,
       ToplamM2,
       ToplamAdet
FROM Siparişler
WHERE (SiparisNo = '{srg}'
       OR Renk IN (SELECT Renk FROM Siparişler WHERE SiparisNo = '{srg}'))
  AND AnaSiparişMi = 'Evet'
ORDER BY CASE 
            WHEN SiparisNo = '{srg}' THEN 0 
            ELSE 1 
         END, 
         Renk ASC, 
         SiparisNo ASC;
";
SqlDataAdapter adap3 = new SqlDataAdapter(sorgu3, bgl.baglanti());
                    DataSet ds3 = new DataSet();
                    adap3.Fill(ds3, "Siparişler");
                    this.dataGridView1.DataSource = ds3.Tables[0];
                    aynısipnugetirme();


                }

                //if (dataGridView1.Rows.Count >= 2)
                //{
                //    string renk = dataGridView1.Rows[0].Cells[3].Value.ToString();
                //    string sorgu3 = "SELECT DISTINCT SiparisNo,Müşteri,Model,Renk,SiparişTipi,SiparişTarihi,KesildiMi,KesildiTarihi,Onay,Palet,Etiket,ToplamM2,ToplamAdet From Siparişler where Renk Like '" + renk + "' AND  AnaSiparişMi='" + "Evet" + "' AND  Aşama='" + "Etiket" + "' ORDER BY SiparisNo ASC";
                //    SqlDataAdapter adap3 = new SqlDataAdapter(sorgu3, bgl.baglanti());
                //    DataSet ds3 = new DataSet();
                //    adap3.Fill(ds3, "Siparişler");
                //    this.dataGridView1.DataSource = ds3.Tables[0];
                //}

                string srg2 = comboBox3.Text;
                string sorgu2 = @"SELECT 
    SiparisNo,
    Müşteri,
    Model,
    Renk,
    SiparişTipi,
    SiparişTarihi,
    KesildiMi,
    KesildiTarihi,
    Onay,
    Palet,
    Etiket,
    ToplamM2,
    ToplamAdet,  
    Boy,
    En,
    Özellik,
    Telefon,
    TeslimTarihi,
    Adres,
    SUM(NULLIF(TRY_CAST(REPLACE(Adet, ',', '.') AS decimal(12, 0)), 0)) AS Adet 
FROM 
    Siparişler 
WHERE 
    SiparisNo LIKE '"+srg2+"' GROUP BY SiparisNo, Müşteri, Model, Renk, Boy, En, SiparişTipi, ToplamAdet, SiparişTarihi, KesildiMi, KesildiTarihi, Onay, Palet, Etiket, ToplamM2, Özellik, Telefon, TeslimTarihi, Adres ORDER BY SiparisNo ASC, Boy DESC;";
                SqlDataAdapter adap2 = new SqlDataAdapter(sorgu2, bgl.baglanti());
                DataSet ds2 = new DataSet();
                adap2.Fill(ds2, "Siparişler");
                this.dataGridView2.DataSource = ds2.Tables[0];
                aynısipnugetirme();
                if (dataGridView1.RowCount > 1 && dataGridView2.RowCount > 1)
                {

                    SqlCommand komut = new SqlCommand();
                    komut.CommandText = "SELECT *FROM Paletler where Renk=@Renk";
                    komut.Parameters.AddWithValue("@Renk", dataGridView2.Rows[0].Cells[3].Value.ToString());
                    komut.Connection = bgl.baglanti();
                    komut.CommandType = CommandType.Text;

                    SqlDataReader dr;
                    dr = komut.ExecuteReader();
                    if (dr.Read())
                    {

                        do
                        {

                            textBox16.Text = dr["Palet"].ToString();
                            textBox16.Enabled = false;
                        } while (dr.Read());
                    }
                    else
                    {
                        textBox16.Text = "";
                        textBox16.Enabled = true;

                    }
                }
            }
            bgl.baglanti().Close();

        }

        string aynıpaletmevcut;

        string siparisno;
        string palett;
        string kesilditarihii;
        string kesildimii;
        string renkk;
        string izin;

        string hepsi;
        string paletlenen;
        string kesilecek;
        string toplamadet;
        private void button10_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                aynıpaletmevcut = "";
                siparisno = "";
                palett = "";
                kesilditarihii = "";
                kesildimii = "";
                renkk = "";
                hepsi = "";
                paletlenen = "";
                kesilecek = "";
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Siparişler WHERE SiparisNo=@SiparisNo AND AnaSiparişMi='" + "Evet" + "' ";
                komut.Parameters.AddWithValue("@SiparisNo", comboBox3.Text);
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    siparisno = dr["SiparisNo"].ToString();
                    if (siparisno == comboBox3.Text)
                    {
                        izin = "var";
                        palett = dr["Palet"].ToString();
                        kesilditarihii = dr["KesildiTarihi"].ToString();
                        kesildimii = dr["KesildiMi"].ToString();
                        renkk = dr["Renk"].ToString();
                        toplamadet = dr["ToplamAdet"].ToString();
                        break;
                    }

                }
                if (/*dataGridView1.Rows.Count == 2 && */izin == "var")
                {
                    if (palett.Length >= 1 && kesilditarihii.Length >= 1)
                    {
                        DialogResult d1 = new DialogResult();
                        d1 = MessageBox.Show("Bu etiket daha önce çıkarılmış. Tekrar çıkartmak istiyor musunuz ?", "Uyarı", MessageBoxButtons.YesNo);
                        if (d1 == DialogResult.Yes)
                        {
                            Excel.Application excel = new Excel.Application();
                            excel.Visible = true;
                            object Missing = Type.Missing;
                            //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Modeks_Dosyalar\\Etiket2.xlsx ");
                            //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Enes\\Desktop\\Etiket2.xlsx");
                            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Modeks_Dosyalar\\Etiket2.xlsx");
                            Excel.Worksheet sheet1 = (Excel.Worksheet)workbook.Sheets[1];

                            int j2 = 2;
                            int x2 = 1;
                            int btr = 0;
                            for (int k = 0; k < dataGridView2.Rows.Count; k++)
                            {
                                int toplamAdet = Convert.ToInt32(dataGridView2.Rows[k].Cells["ToplamAdet"].Value);
                                int adet = Convert.ToInt32(dataGridView2.Rows[k].Cells["Adet"].Value);

                                for (int i = 0; i < adet; i++)
                                {
                                    sheet1.Cells[j2, 2].Value = dataGridView2.Rows[k].Cells[0].Value.ToString();
                                    sheet1.Cells[j2, 4].Value = toplamAdet;
                                    sheet1.Cells[3 + j2, 4].Value = adet.ToString();
                                    sheet1.Cells[5 + j2, 4].Value = dataGridView2.Rows[k].Cells[13].Value.ToString();
                                    sheet1.Cells[7 + j2, 4].Value = dataGridView2.Rows[k].Cells[14].Value.ToString();
                                    sheet1.Cells[4 + j2, 2].Value = dataGridView2.Rows[k].Cells[15].Value.ToString();
                                    sheet1.Cells[j2 - 1, 2].Value = dataGridView2.Rows[k].Cells[1].Value.ToString();
                                    sheet1.Cells[2 + j2, 2].Value = dataGridView2.Rows[k].Cells[3].Value.ToString();
                                    sheet1.Cells[3 + j2, 2].Value = dataGridView2.Rows[k].Cells[2].Value.ToString();

                                    sheet1.Cells[(7 + j2), 2].Value = $"{x2}/{adet}";
                                    btr++;
                                    if (btr != toplamAdet)
                                    {
                                        sheet1.Range["A1:F10"].Copy(sheet1.Range["A" + (j2 + 9) + ""]);
                                    }

                                    j2 += 10;

                                    if (x2 == adet)
                                    {
                                        x2 = 1; 
                                    }
                                    else
                                    {
                                        x2++;
                                    }
                                }
                            }


                            string etiketFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Etiket");

                            if (!Directory.Exists(etiketFolderPath))
                            {
                                Directory.CreateDirectory(etiketFolderPath);
                            }

                            string excelFilePath = Path.Combine(etiketFolderPath, "Etiket" + siparisno + ".xlsx");

                            try
                            {
                                workbook.SaveAs(excelFilePath);

                                sheet1.PrintOutEx(Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing);
                                workbook.Close(false, Missing, Missing);
                                excel.Quit();
                                foreach (var process in Process.GetProcessesByName("EXCEL"))
                                {
                                    process.Kill();
                                }
                            }
                            catch (Exception)
                            {
                                MessageBox.Show("Etiketi kaydederken bir hata oluştu, muhtemelen seçeneklerde 'hayır' butonuna bastınız, tekrar deneyiniz..");
                            }
                            
                            string tarih = Convert.ToDateTime(DateTime.Now).ToString("yyyy-MM-dd HH:mm:ss");

                            string sorgu3 = "UPDATE Siparişler SET Etiket=@Etiket WHERE SiparisNo=@SiparisNo AND AnaSiparişMi='" + "Evet" + "' ";
                            SqlCommand komut3;
                            komut3 = new SqlCommand(sorgu3, bgl.baglanti());
                            komut3.Parameters.AddWithValue("@SiparisNo", comboBox3.Text);
                            komut3.Parameters.AddWithValue("@Etiket", tarih);
                            komut3.ExecuteNonQuery();
                            label30.Text = comboBox3.Text;
                            label29.Text = dataGridView1.Rows[0].Cells["Müşteri"].Value.ToString();
                            Methodlar();
                        }


                    }
                    else if (palett.Length >= 1 && kesilditarihii.Length <= 1)
                    {
                      
                        if (kesildimii == "Evet")
                        {
                            Excel.Application excel = new Excel.Application();
                            excel.Visible = true;
                            object Missing = Type.Missing;
                            //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Modeks_Dosyalar\\Etiket2.xlsx ");
                            //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Enes\\Desktop\\Etiket2.xlsx");
                            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Modeks_Dosyalar\\Etiket2.xlsx");
                            Excel.Worksheet sheet1 = (Excel.Worksheet)workbook.Sheets[1];
                            int j2 = 2;
                            int x2 = 1;
                            int btr = 0;
                            for (int k = 0; k < dataGridView2.Rows.Count; k++)
                            {
                                int toplamAdet = Convert.ToInt32(dataGridView2.Rows[k].Cells["ToplamAdet"].Value);
                                int adet = Convert.ToInt32(dataGridView2.Rows[k].Cells["Adet"].Value);

                                for (int i = 0; i < adet; i++)
                                {
                                    sheet1.Cells[j2, 2].Value = dataGridView2.Rows[k].Cells[0].Value.ToString();
                                    sheet1.Cells[j2, 4].Value = toplamAdet;
                                    sheet1.Cells[3 + j2, 4].Value = adet.ToString();
                                    sheet1.Cells[5 + j2, 4].Value = dataGridView2.Rows[k].Cells[13].Value.ToString();
                                    sheet1.Cells[7 + j2, 4].Value = dataGridView2.Rows[k].Cells[14].Value.ToString();
                                    sheet1.Cells[4 + j2, 2].Value = dataGridView2.Rows[k].Cells[15].Value.ToString();
                                    sheet1.Cells[j2 - 1, 2].Value = dataGridView2.Rows[k].Cells[1].Value.ToString();
                                    sheet1.Cells[2 + j2, 2].Value = dataGridView2.Rows[k].Cells[3].Value.ToString();
                                    sheet1.Cells[3 + j2, 2].Value = dataGridView2.Rows[k].Cells[2].Value.ToString();

                                    sheet1.Cells[(7 + j2), 2].Value = $"{x2}/{adet}";
                                    btr++;
                                    if (btr != toplamAdet)
                                    {
                                        sheet1.Range["A1:F10"].Copy(sheet1.Range["A" + (j2 + 9) + ""]);
                                    }

                                    j2 += 10;

                                    if (x2 == adet)
                                    {
                                        x2 = 1;
                                    }
                                    else
                                    {
                                        x2++;
                                    }
                                }
                            }
                            string etiketFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Etiket");

                            if (!Directory.Exists(etiketFolderPath))
                            {
                                Directory.CreateDirectory(etiketFolderPath);
                            }

                            string excelFilePath = Path.Combine(etiketFolderPath, "Etiket" + siparisno + ".xlsx");


                            workbook.SaveAs(excelFilePath);
                            //sheet1.Range["A14:F23"].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                            sheet1.PrintOutEx(Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing);

                            // Excel dosyasını kapatma
                            workbook.Close(false, Missing, Missing);
                            excel.Quit();
                            foreach (var process in Process.GetProcessesByName("EXCEL"))
                            {
                                process.Kill();
                            }
                            //sheet1.PrintPreview();
                            //workbook.Close(false);
                            //excel.Quit();

                            string tarih = Convert.ToDateTime(DateTime.Now).ToString("yyyy-MM-dd HH:mm:ss");

                            string sorgu3 = "UPDATE Siparişler SET Etiket=@Etiket WHERE SiparisNo=@SiparisNo AND AnaSiparişMi='" + "Evet" + "' ";
                            SqlCommand komut3;
                            komut3 = new SqlCommand(sorgu3, bgl.baglanti());
                            komut3.Parameters.AddWithValue("@SiparisNo", comboBox3.Text);
                            komut3.Parameters.AddWithValue("@Etiket", tarih);
                            komut3.ExecuteNonQuery();
                            label30.Text = comboBox3.Text;
                            label29.Text = dataGridView1.Rows[0].Cells["Müşteri"].Value.ToString();
                            Methodlar();
                            comboBox3.Text = "";
                            textBox17.Text = "";
                        }
                        else
                        {
                            MessageBox.Show("Önce siparişi Cnc'de kesiniz!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        if (textBox16.Text != "")
                        {
                            SqlCommand komut3 = new SqlCommand();
                            komut3.CommandText = "SELECT *FROM Paletler";
                            komut3.Connection = bgl.baglanti();
                            komut3.CommandType = CommandType.Text;

                            SqlDataReader dr3;
                            dr3 = komut3.ExecuteReader();
                            while (dr3.Read())
                            {
                                if (textBox16.Text == dr3["Palet"].ToString())
                                {
                                    if (dr3["Renk"].ToString() != renkk)
                                    {
                                        aynıpaletmevcut = "evet";
                                        break;
                                    }
                                }

                            }
                            if (aynıpaletmevcut != "evet")
                            {
                                if (kesildimii == "Evet")
                                {
                                    string sorgu2 = "UPDATE Siparişler SET Palet=@Palet WHERE SiparisNo=@SiparisNo AND AnaSiparişMi='" + "Evet" + "' ";
                                    SqlCommand komut2;
                                    komut2 = new SqlCommand(sorgu2, bgl.baglanti());
                                    komut2.Parameters.AddWithValue("@SiparisNo", comboBox3.Text);
                                    komut2.Parameters.AddWithValue("@Palet", textBox16.Text);
                                    MessageBox.Show(comboBox3.Text + " lu sipariş " + textBox16.Text + " palete yerleştirilmiştir.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    komut2.ExecuteNonQuery();


                                    SqlCommand komut5 = new SqlCommand();
                                    komut5.CommandText = "SELECT *From Grafik where Renk=@Renk";
                                    komut5.Parameters.AddWithValue("@Renk", dataGridView1.Rows[0].Cells["Renk"].Value.ToString());
                                    komut5.Connection = bgl.baglanti();
                                    komut5.CommandType = CommandType.Text;
                                    SqlDataReader dr2;
                                    dr2 = komut5.ExecuteReader();
                                    while (dr2.Read())
                                    {
                                        hepsi = dr2["Hepsi"].ToString();
                                        paletlenen = dr2["Paletlenen"].ToString();
                                        kesilecek = dr2["Paletlenecek"].ToString();
                                    }


                                    string sorgu4 = "UPDATE Grafik SET Paletlenecek=@Paletlenecek, Paletlenen=@Paletlenen, Palet=@Palet WHERE Renk=@Renk";
                                    SqlCommand komut4;
                                    komut4 = new SqlCommand(sorgu4, bgl.baglanti());
                                    komut4.Parameters.AddWithValue("@Renk", dataGridView1.Rows[0].Cells["Renk"].Value.ToString());
                                    komut4.Parameters.AddWithValue("@Paletlenecek", Convert.ToDouble(Convert.ToDouble(kesilecek) - Convert.ToDouble(dataGridView1.Rows[0].Cells["ToplamM2"].Value.ToString())));
                                    komut4.Parameters.AddWithValue("@Paletlenen", Convert.ToDouble(Convert.ToDouble(paletlenen) + Convert.ToDouble(dataGridView1.Rows[0].Cells["ToplamM2"].Value.ToString())));
                                    komut4.Parameters.AddWithValue("@Palet", textBox16.Text);
                                    komut4.ExecuteNonQuery();


                                    string kayit = "insert into Paletler(Renk,Palet)values (@p1,@p2)";
                                    SqlCommand cmd = new SqlCommand(kayit, bgl.baglanti());
                                    cmd.Parameters.AddWithValue("@p1", renkk);
                                    cmd.Parameters.AddWithValue("@p2", textBox16.Text);
                                    cmd.ExecuteNonQuery();
                                }
                                else
                                {
                                    MessageBox.Show("Önce siparişi Cnc'de kesiniz! Onaylı değil ise onaylayınız.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }



                                int j = 2;
                                int x = 1;
                                //if (dataGridView1.Rows.Count == 2)
                                //{
                                //for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
                                //{
                                //    if (dataGridView1.Rows[i].Visible == true)
                                //        palet = textBox16.Text;
                                //}
                                //if (palett.Length >= 1)
                                //{
                                if (kesildimii == "Evet")
                                {
                                    Excel.Application excel = new Excel.Application();
                                    excel.Visible = true;
                                    object Missing = Type.Missing;
                                    //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Modeks_Dosyalar\\Etiket2.xlsx ");
                                    //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Enes\\Desktop\\Etiket2.xlsx");
                                    Excel.Workbook workbook = excel.Workbooks.Open("C:\\Modeks_Dosyalar\\Etiket2.xlsx");
                                    Excel.Worksheet sheet1 = (Excel.Worksheet)workbook.Sheets[1];

                                    int j2 = 2;
                                    int x2 = 1;
                                    int btr = 0;
                                    for (int k = 0; k < dataGridView2.Rows.Count; k++)
                                    {
                                        int toplamAdet = Convert.ToInt32(dataGridView2.Rows[k].Cells["ToplamAdet"].Value);
                                        int adet = Convert.ToInt32(dataGridView2.Rows[k].Cells["Adet"].Value);

                                        for (int i = 0; i < adet; i++)
                                        {
                                            sheet1.Cells[j2, 2].Value = dataGridView2.Rows[k].Cells[0].Value.ToString();
                                            sheet1.Cells[j2, 4].Value = toplamAdet;
                                            sheet1.Cells[3 + j2, 4].Value = adet.ToString();
                                            sheet1.Cells[5 + j2, 4].Value = dataGridView2.Rows[k].Cells[13].Value.ToString();
                                            sheet1.Cells[7 + j2, 4].Value = dataGridView2.Rows[k].Cells[14].Value.ToString();
                                            sheet1.Cells[4 + j2, 2].Value = dataGridView2.Rows[k].Cells[15].Value.ToString();
                                            sheet1.Cells[j2 - 1, 2].Value = dataGridView2.Rows[k].Cells[1].Value.ToString();
                                            sheet1.Cells[2 + j2, 2].Value = dataGridView2.Rows[k].Cells[3].Value.ToString();
                                            sheet1.Cells[3 + j2, 2].Value = dataGridView2.Rows[k].Cells[2].Value.ToString();

                                            sheet1.Cells[(7 + j2), 2].Value = $"{x2}/{adet}";
                                            btr++;
                                            if (btr != toplamAdet)
                                            {
                                                sheet1.Range["A1:F10"].Copy(sheet1.Range["A" + (j2 + 9) + ""]);
                                            }

                                            j2 += 10;

                                            if (x2 == adet)
                                            {
                                                x2 = 1;
                                            }
                                            else
                                            {
                                                x2++;
                                            }
                                        }
                                    }
                                    string etiketFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Etiket");

                                    if (!Directory.Exists(etiketFolderPath))
                                    {
                                        Directory.CreateDirectory(etiketFolderPath);
                                    }

                                    string excelFilePath = Path.Combine(etiketFolderPath, "Etiket" + siparisno + ".xlsx");


                                    workbook.SaveAs(excelFilePath);
                                    //sheet1.Range["A14:F23"].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                                    sheet1.PrintOutEx(Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing);

                                    // Excel dosyasını kapatma
                                    workbook.Close(false, Missing, Missing);
                                    excel.Quit();
                                    foreach (var process in Process.GetProcessesByName("EXCEL"))
                                    {
                                        process.Kill();
                                    }
                                    string tarih = Convert.ToDateTime(DateTime.Now).ToString("yyyy-MM-dd HH:mm:ss");

                                    string sorgu4 = "UPDATE Siparişler SET Etiket=@Etiket WHERE SiparisNo=@SiparisNo AND AnaSiparişMi='" + "Evet" + "' ";
                                    SqlCommand komut4;
                                    komut4 = new SqlCommand(sorgu4, bgl.baglanti());
                                    komut4.Parameters.AddWithValue("@SiparisNo", comboBox3.Text);
                                    komut4.Parameters.AddWithValue("@Etiket", tarih);
                                    komut4.ExecuteNonQuery();
                                    label30.Text = comboBox3.Text;
                                    label29.Text = dataGridView1.Rows[0].Cells["Müşteri"].Value.ToString();
                                    liste();
                                    liste2();
                                    Methodlar();
                                    comboBox3.Text = "";
                                    textBox17.Text = "";
                                }
                                else
                                {
                                    //MessageBox.Show("Önce siparişi Cnc'de kesiniz!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                                //}
                                //else
                                //{
                                //    MessageBox.Show("Önce siparişe bir palet numarası verin!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                //}
                                //}
                                //else
                                //{
                                //    MessageBox.Show("Bir sipariş numarası giriniz!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                //}
                            }
                            else
                            {
                                MessageBox.Show("Bu palet başka bir renge verilmiş! Başka bir palet numarası giriniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }

                        }
                        else
                        {
                            MessageBox.Show("Bir palet numarası giriniz!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                    }

                }
                else
                {
                    MessageBox.Show("Bir sipariş numarası giriniz!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                textBox16.Text = "";
            }
            else if (radioButton2.Checked == true)
            {
                aynıpaletmevcut = "";
                siparisno = "";
                palett = "";
                kesilditarihii = "";
                kesildimii = "";
                renkk = "";
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Siparişler WHERE SiparisNo=@SiparisNo AND AnaSiparişMi='" + "Evet" + "' ";
                komut.Parameters.AddWithValue("@SiparisNo", comboBox3.Text);
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    siparisno = dr["SiparisNo"].ToString();
                    if (siparisno == comboBox3.Text)
                    {
                        izin = "var";
                        palett = dr["Palet"].ToString();
                        kesilditarihii = dr["KesildiTarihi"].ToString();
                        kesildimii = dr["KesildiMi"].ToString();
                        renkk = dr["Renk"].ToString();
                        toplamadet = dr["ToplamAdet"].ToString();
                        break;
                    }

                }
                if (/*dataGridView1.Rows.Count == 2 && */izin == "var")
                {
                    if (palett.Length >= 1 && kesilditarihii.Length >= 1)
                    {
                        DialogResult d1 = new DialogResult();
                        d1 = MessageBox.Show("Bu etiket daha önce çıkarılmış. Tekrar çıkartmak istiyor musunuz ?", "Uyarı", MessageBoxButtons.YesNo);
                        if (d1 == DialogResult.Yes)
                        {
                            Excel.Application excel = new Excel.Application();
                            excel.Visible = true;
                            object Missing = Type.Missing;
                            //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Modeks_Dosyalar\\Etiket2Logosuz.xlsx ");
                            //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Enes\\Desktop\\Etiket2Logosuz.xlsx");
                            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Modeks_Dosyalar\\Etiket2Logosuz.xlsx");
                            Excel.Worksheet sheet1 = (Excel.Worksheet)workbook.Sheets[1];
                            int j2 = 2;
                            int x2 = 1;
                            int btr = 0;
                            for (int k = 0; k < dataGridView2.Rows.Count; k++)
                            {
                                int toplamAdet = Convert.ToInt32(dataGridView2.Rows[k].Cells["ToplamAdet"].Value);
                                int adet = Convert.ToInt32(dataGridView2.Rows[k].Cells["Adet"].Value);

                                for (int i = 0; i < adet; i++)
                                {
                                    sheet1.Cells[j2, 2].Value = dataGridView2.Rows[k].Cells[0].Value.ToString();
                                    sheet1.Cells[j2, 4].Value = toplamAdet;
                                    sheet1.Cells[3 + j2, 4].Value = adet.ToString();
                                    sheet1.Cells[5 + j2, 4].Value = dataGridView2.Rows[k].Cells[13].Value.ToString();
                                    sheet1.Cells[7 + j2, 4].Value = dataGridView2.Rows[k].Cells[14].Value.ToString();
                                    sheet1.Cells[4 + j2, 2].Value = dataGridView2.Rows[k].Cells[15].Value.ToString();
                                    sheet1.Cells[j2 - 1, 2].Value = dataGridView2.Rows[k].Cells[1].Value.ToString();
                                    sheet1.Cells[2 + j2, 2].Value = dataGridView2.Rows[k].Cells[3].Value.ToString();
                                    sheet1.Cells[3 + j2, 2].Value = dataGridView2.Rows[k].Cells[2].Value.ToString();

                                    sheet1.Cells[(7 + j2), 2].Value = $"{x2}/{adet}";
                                    btr++;
                                    if (btr != toplamAdet)
                                    {
                                        sheet1.Range["A1:F10"].Copy(sheet1.Range["A" + (j2 + 9) + ""]);
                                    }

                                    j2 += 10;

                                    if (x2 == adet)
                                    {
                                        x2 = 1;
                                    }
                                    else
                                    {
                                        x2++;
                                    }
                                }
                            }
                            string etiketFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Etiket");

                            if (!Directory.Exists(etiketFolderPath))
                            {
                                Directory.CreateDirectory(etiketFolderPath);
                            }

                            string excelFilePath = Path.Combine(etiketFolderPath, "Etiket" + siparisno + ".xlsx");


                            workbook.SaveAs(excelFilePath);
                            //sheet1.Range["A14:F23"].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                            sheet1.PrintOutEx(Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing);
                            workbook.Close(false, Missing, Missing);
                            excel.Quit();
                            foreach (var process in Process.GetProcessesByName("EXCEL"))
                            {
                                process.Kill();
                            }
                        }

                    }
                    else if (palett.Length >= 1 && kesilditarihii.Length <= 1)
                    {
                        //if (dataGridView1.Rows.Count == 2)
                        //{
                        //for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
                        //{
                        //    if (dataGridView1.Rows[i].Visible == true)
                        //        palet = dataGridView1.Rows[0].Cells[9].Value.ToString();
                        //}
                        //if (palett.Length >= 1)
                        //{
                        if (kesildimii == "Evet")
                        {
                            Excel.Application excel = new Excel.Application();
                            excel.Visible = true;
                            object Missing = Type.Missing;
                            //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Modeks_Dosyalar\\Etiket2Logosuz.xlsx ");
                            //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Enes\\Desktop\\Etiket2Logosuz.xlsx");
                            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Modeks_Dosyalar\\Etiket2Logosuz.xlsx");
                            Excel.Worksheet sheet1 = (Excel.Worksheet)workbook.Sheets[1];
                            int j2 = 2;
                            int x2 = 1;
                            int btr = 0;
                            for (int k = 0; k < dataGridView2.Rows.Count; k++)
                            {
                                int toplamAdet = Convert.ToInt32(dataGridView2.Rows[k].Cells["ToplamAdet"].Value);
                                int adet = Convert.ToInt32(dataGridView2.Rows[k].Cells["Adet"].Value);

                                for (int i = 0; i < adet; i++)
                                {
                                    sheet1.Cells[j2, 2].Value = dataGridView2.Rows[k].Cells[0].Value.ToString();
                                    sheet1.Cells[j2, 4].Value = toplamAdet;
                                    sheet1.Cells[3 + j2, 4].Value = adet.ToString();
                                    sheet1.Cells[5 + j2, 4].Value = dataGridView2.Rows[k].Cells[13].Value.ToString();
                                    sheet1.Cells[7 + j2, 4].Value = dataGridView2.Rows[k].Cells[14].Value.ToString();
                                    sheet1.Cells[4 + j2, 2].Value = dataGridView2.Rows[k].Cells[15].Value.ToString();
                                    sheet1.Cells[j2 - 1, 2].Value = dataGridView2.Rows[k].Cells[1].Value.ToString();
                                    sheet1.Cells[2 + j2, 2].Value = dataGridView2.Rows[k].Cells[3].Value.ToString();
                                    sheet1.Cells[3 + j2, 2].Value = dataGridView2.Rows[k].Cells[2].Value.ToString();

                                    sheet1.Cells[(7 + j2), 2].Value = $"{x2}/{adet}";
                                    btr++;
                                    if (btr != toplamAdet)
                                    {
                                        sheet1.Range["A1:F10"].Copy(sheet1.Range["A" + (j2 + 9) + ""]);
                                    }

                                    j2 += 10;

                                    if (x2 == adet)
                                    {
                                        x2 = 1;
                                    }
                                    else
                                    {
                                        x2++;
                                    }
                                }
                            }
                            string etiketFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Etiket");

                            if (!Directory.Exists(etiketFolderPath))
                            {
                                Directory.CreateDirectory(etiketFolderPath);
                            }

                            string excelFilePath = Path.Combine(etiketFolderPath, "Etiket" + siparisno + ".xlsx");


                            workbook.SaveAs(excelFilePath);
                            //sheet1.Range["A14:F23"].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                            sheet1.PrintOutEx(Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing);
                            workbook.Close(false, Missing, Missing);
                            excel.Quit();
                            foreach (var process in Process.GetProcessesByName("EXCEL"))
                            {
                                process.Kill();
                            }
                            string tarih = Convert.ToDateTime(DateTime.Now).ToString("yyyy-MM-dd HH:mm:ss");

                            string sorgu3 = "UPDATE Siparişler SET Etiket=@Etiket WHERE SiparisNo=@SiparisNo AND AnaSiparişMi='" + "Evet" + "' ";
                            SqlCommand komut3;
                            komut3 = new SqlCommand(sorgu3, bgl.baglanti());
                            komut3.Parameters.AddWithValue("@SiparisNo", comboBox3.Text);
                            komut3.Parameters.AddWithValue("@Etiket", tarih);
                            komut3.ExecuteNonQuery();
                            label30.Text = comboBox3.Text;
                            label29.Text = dataGridView1.Rows[0].Cells["Müşteri"].Value.ToString();
                            Methodlar();
                            comboBox3.Text = "";
                            textBox17.Text = "";
                        }
                        else
                        {
                            MessageBox.Show("Önce siparişi Cnc'de kesiniz!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        //}
                        //else
                        //{
                        //    MessageBox.Show("Önce siparişe bir palet numarası verin!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        //}
                        //}
                    }
                    else
                    {
                        if (textBox16.Text != "")
                        {
                            SqlCommand komut3 = new SqlCommand();
                            komut3.CommandText = "SELECT *FROM Paletler";
                            komut3.Connection = bgl.baglanti();
                            komut3.CommandType = CommandType.Text;

                            SqlDataReader dr3;
                            dr3 = komut3.ExecuteReader();
                            while (dr3.Read())
                            {
                                if (textBox16.Text == dr3["Palet"].ToString())
                                {
                                    if (dr3["Renk"].ToString() != renkk)
                                    {
                                        aynıpaletmevcut = "evet";
                                        break;
                                    }
                                }

                            }
                            if (aynıpaletmevcut != "evet")
                            {
                                if (kesildimii == "Evet")
                                {
                                    string sorgu2 = "UPDATE Siparişler SET Palet=@Palet WHERE SiparisNo=@SiparisNo AND AnaSiparişMi='" + "Evet" + "' ";
                                    SqlCommand komut2;
                                    komut2 = new SqlCommand(sorgu2, bgl.baglanti());
                                    komut2.Parameters.AddWithValue("@SiparisNo", comboBox3.Text);
                                    komut2.Parameters.AddWithValue("@Palet", textBox16.Text);
                                    MessageBox.Show(comboBox3.Text + " lu sipariş " + textBox16.Text + " palete yerleştirilmiştir.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    komut2.ExecuteNonQuery();

                                    SqlCommand komut5 = new SqlCommand();
                                    komut5.CommandText = "SELECT *From Grafik where Renk=@Renk";
                                    komut5.Parameters.AddWithValue("@Renk", dataGridView1.Rows[0].Cells["Renk"].Value.ToString());
                                    komut5.Connection = bgl.baglanti();
                                    komut5.CommandType = CommandType.Text;
                                    SqlDataReader dr2;
                                    dr2 = komut5.ExecuteReader();
                                    while (dr2.Read())
                                    {
                                        hepsi = dr2["Hepsi"].ToString();
                                        paletlenen = dr2["Paletlenen"].ToString();
                                        kesilecek = dr2["Paletlenecek"].ToString();
                                    }


                                    string sorgu4 = "UPDATE Grafik SET Paletlenecek=@Paletlenecek, Paletlenen=@Paletlenen, Palet=@Palet WHERE Renk=@Renk";
                                    SqlCommand komut4;
                                    komut4 = new SqlCommand(sorgu4, bgl.baglanti());
                                    komut4.Parameters.AddWithValue("@Renk", dataGridView1.Rows[0].Cells["Renk"].Value.ToString());
                                    komut4.Parameters.AddWithValue("@Paletlenecek", Convert.ToDouble(Convert.ToDouble(kesilecek) - Convert.ToDouble(dataGridView1.Rows[0].Cells["ToplamM2"].Value.ToString())));
                                    komut4.Parameters.AddWithValue("@Paletlenen", Convert.ToDouble(Convert.ToDouble(paletlenen) + Convert.ToDouble(dataGridView1.Rows[0].Cells["ToplamM2"].Value.ToString())));
                                    komut4.Parameters.AddWithValue("@Palet", textBox16.Text);
                                    komut4.ExecuteNonQuery();


                                    string kayit = "insert into Paletler(Renk,Palet)values (@p1,@p2)";
                                    SqlCommand cmd = new SqlCommand(kayit, bgl.baglanti());
                                    cmd.Parameters.AddWithValue("@p1", renkk);
                                    cmd.Parameters.AddWithValue("@p2", textBox16.Text);
                                    cmd.ExecuteNonQuery();
                                }
                                else
                                {
                                    MessageBox.Show("Önce siparişi Cnc'de kesiniz! Onaylı değil ise onaylayınız.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }

                                //if (dataGridView1.Rows.Count == 2)
                                //{
                                //for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
                                //{
                                //    if (dataGridView1.Rows[i].Visible == true)
                                //        palet = textBox16.Text;
                                //}
                                //if (palett.Length >= 1)
                                //{
                                if (kesildimii == "Evet")
                                {
                                    Excel.Application excel = new Excel.Application();
                                    excel.Visible = true;
                                    object Missing = Type.Missing;
                                    //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Modeks_Dosyalar\\Etiket2Logosuz.xlsx ");
                                    //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Enes\\Desktop\\Etiket2Logosuz.xlsx");
                                    Excel.Workbook workbook = excel.Workbooks.Open("C:\\Modeks_Dosyalar\\Etiket2Logosuz.xlsx");
                                    Excel.Worksheet sheet1 = (Excel.Worksheet)workbook.Sheets[1];

                                    int j2 = 2;
                                    int x2 = 1;
                                    int btr = 0;
                                    for (int k = 0; k < dataGridView2.Rows.Count; k++)
                                    {
                                        int toplamAdet = Convert.ToInt32(dataGridView2.Rows[k].Cells["ToplamAdet"].Value);
                                        int adet = Convert.ToInt32(dataGridView2.Rows[k].Cells["Adet"].Value);

                                        for (int i = 0; i < adet; i++)
                                        {
                                            sheet1.Cells[j2, 2].Value = dataGridView2.Rows[k].Cells[0].Value.ToString();
                                            sheet1.Cells[j2, 4].Value = toplamAdet;
                                            sheet1.Cells[3 + j2, 4].Value = adet.ToString();
                                            sheet1.Cells[5 + j2, 4].Value = dataGridView2.Rows[k].Cells[13].Value.ToString();
                                            sheet1.Cells[7 + j2, 4].Value = dataGridView2.Rows[k].Cells[14].Value.ToString();
                                            sheet1.Cells[4 + j2, 2].Value = dataGridView2.Rows[k].Cells[15].Value.ToString();
                                            sheet1.Cells[j2 - 1, 2].Value = dataGridView2.Rows[k].Cells[1].Value.ToString();
                                            sheet1.Cells[2 + j2, 2].Value = dataGridView2.Rows[k].Cells[3].Value.ToString();
                                            sheet1.Cells[3 + j2, 2].Value = dataGridView2.Rows[k].Cells[2].Value.ToString();

                                            sheet1.Cells[(7 + j2), 2].Value = $"{x2}/{adet}";
                                            btr++;
                                            if (btr != toplamAdet)
                                            {
                                                sheet1.Range["A1:F10"].Copy(sheet1.Range["A" + (j2 + 9) + ""]);
                                            }

                                            j2 += 10;

                                            if (x2 == adet)
                                            {
                                                x2 = 1;
                                            }
                                            else
                                            {
                                                x2++;
                                            }
                                        }
                                    }
                                    string etiketFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Etiket");

                                    if (!Directory.Exists(etiketFolderPath))
                                    {
                                        Directory.CreateDirectory(etiketFolderPath);
                                    }

                                    string excelFilePath = Path.Combine(etiketFolderPath, "Etiket" + siparisno + ".xlsx");

                                    workbook.SaveAs(excelFilePath);
                                    //sheet1.Range["A14:F23"].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                                    sheet1.PrintOutEx(Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing);
                                    workbook.Close(false, Missing, Missing);
                                    excel.Quit();
                                    foreach (var process in Process.GetProcessesByName("EXCEL"))
                                    {
                                        process.Kill();
                                    }
                                    string tarih = Convert.ToDateTime(DateTime.Now).ToString("yyyy-MM-dd HH:mm:ss");

                                    string sorgu4 = "UPDATE Siparişler SET Etiket=@Etiket WHERE SiparisNo=@SiparisNo AND AnaSiparişMi='" + "Evet" + "' ";
                                    SqlCommand komut4;
                                    komut4 = new SqlCommand(sorgu4, bgl.baglanti());
                                    komut4.Parameters.AddWithValue("@SiparisNo", comboBox3.Text);
                                    komut4.Parameters.AddWithValue("@Etiket", tarih);
                                    komut4.ExecuteNonQuery();
                                    label30.Text = comboBox3.Text;
                                    label29.Text = dataGridView1.Rows[0].Cells["Müşteri"].Value.ToString();
                                    liste();
                                    liste2();
                                    Methodlar();
                                    comboBox3.Text = "";
                                    textBox17.Text = "";
                                }
                                else
                                {
                                    //MessageBox.Show("Önce siparişi Cnc'de kesiniz!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                                //}
                                //else
                                //{
                                //    MessageBox.Show("Önce siparişe bir palet numarası verin!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                //}
                                //}
                                //else
                                //{
                                //    MessageBox.Show("Bir sipariş numarası giriniz!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                //}
                            }
                            else
                            {
                                MessageBox.Show("Bu palet başka bir renge verilmiş! Başka bir palet numarası giriniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }

                        }
                        else
                        {
                            MessageBox.Show("Bir palet numarası giriniz!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                    }

                }
                else
                {
                    MessageBox.Show("Bir sipariş numarası giriniz!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                textBox16.Text = "";
            }
            else if(radioButton3.Checked == true)
            {
                siparisno = "";
                palett = "";
                kesilditarihii = "";
                kesildimii = "";
                renkk = "";
                hepsi = "";
                paletlenen = "";
                kesilecek = "";
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Siparişler WHERE SiparisNo=@SiparisNo AND AnaSiparişMi='" + "Evet" + "' ";
                komut.Parameters.AddWithValue("@SiparisNo", comboBox3.Text);
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    siparisno = dr["SiparisNo"].ToString();
                    if (siparisno == comboBox3.Text)
                    {
                        izin = "var";
                        palett = dr["Palet"].ToString();
                        kesilditarihii = dr["KesildiTarihi"].ToString();
                        kesildimii = dr["KesildiMi"].ToString();
                        renkk = dr["Renk"].ToString();
                        toplamadet = dr["ToplamAdet"].ToString();
                        break;
                    }

                }
               
                        Excel.Application excel = new Excel.Application();
                        excel.Visible = true;
                        object Missing = Type.Missing;
                        Excel.Workbook workbook = excel.Workbooks.Open("C:\\Modeks_Dosyalar\\Etiket2-Yatay.xlsx");
                        Excel.Worksheet sheet1 = (Excel.Worksheet)workbook.Sheets[1];

                        int j2 = 2;
                        int x2 = 1;
                        int btr = 0;
                        for (int k = 0; k < dataGridView2.Rows.Count; k++)
                        {
                            int toplamAdet = Convert.ToInt32(dataGridView2.Rows[k].Cells["ToplamAdet"].Value);
                            int adet = Convert.ToInt32(dataGridView2.Rows[k].Cells["Adet"].Value);

                            for (int i = 0; i < adet; i++)
                            {
                                // satır - sütun

                                sheet1.Cells[j2 - 1, 5].Value = dataGridView2.Rows[k].Cells[0].Value.ToString(); // Sipariş Numarası
                                sheet1.Cells[j2 - 1, 4].Value = dataGridView2.Rows[k].Cells[2].Value.ToString(); // Model
                                sheet1.Cells[j2 - 1, 7].Value = toplamAdet;
                                sheet1.Cells[j2 + 1, 1].Value = dataGridView2.Rows[k].Cells[1].Value.ToString(); // Müşteri Adı
                                sheet1.Cells[j2 + 2, 1].Value = dataGridView2.Rows[k].Cells[3].Value.ToString(); // Renk
                                sheet1.Cells[j2 + 4, 6].Value = adet.ToString();
                                sheet1.Cells[j2 + 5, 4].Value = dataGridView2.Rows[k].Cells[13].Value.ToString();
                                sheet1.Cells[j2 + 7, 4].Value = dataGridView2.Rows[k].Cells[14].Value.ToString();
                                sheet1.Cells[j2 + 4, 1].Value = dataGridView2.Rows[k].Cells[15].Value.ToString();

                                sheet1.Cells[(7 + j2), 1].Value = $"{x2}/{adet}"; 
                        btr++;
                                if (btr != toplamAdet)
                                {
                                    sheet1.Range["A1:H10"].Copy(sheet1.Range["A" + (j2 + 9) + ""]);
                                }

                                j2 += 10;

                                if (x2 == adet)
                                {
                                    x2 = 1;
                                }
                                else
                                {
                                    x2++;
                                }
                            }
                        }


                        string etiketFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Etiket");

                        if (!Directory.Exists(etiketFolderPath))
                        {
                            Directory.CreateDirectory(etiketFolderPath);
                        }

                        string excelFilePath = Path.Combine(etiketFolderPath, "Etiket" + siparisno + ".xlsx");

                        try
                        {
                            workbook.SaveAs(excelFilePath);

                            sheet1.PrintOutEx(Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing);
                            workbook.Close(false, Missing, Missing);
                            excel.Quit();
                            foreach (var process in Process.GetProcessesByName("EXCEL"))
                            {
                                process.Kill();
                            }
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("Etiketi kaydederken bir hata oluştu, muhtemelen seçeneklerde 'hayır' butonuna bastınız, tekrar deneyiniz..");
                        }

                        string tarih = Convert.ToDateTime(DateTime.Now).ToString("yyyy-MM-dd HH:mm:ss");

                        string sorgu3 = "UPDATE Siparişler SET Etiket=@Etiket WHERE SiparisNo=@SiparisNo AND AnaSiparişMi='" + "Evet" + "' ";
                        SqlCommand komut3;
                        komut3 = new SqlCommand(sorgu3, bgl.baglanti());
                        komut3.Parameters.AddWithValue("@SiparisNo", comboBox3.Text);
                        komut3.Parameters.AddWithValue("@Etiket", tarih);
                        komut3.ExecuteNonQuery();
                        label30.Text = comboBox3.Text;
                        label29.Text = dataGridView1.Rows[0].Cells["Müşteri"].Value.ToString();
                        Methodlar();

                textBox16.Text = "";
            } else if(radioButton4.Checked == true)
            {
                siparisno = "";
                palett = "";
                kesilditarihii = "";
                kesildimii = "";
                renkk = "";
                hepsi = "";
                paletlenen = "";
                kesilecek = "";
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Siparişler WHERE SiparisNo=@SiparisNo AND AnaSiparişMi='" + "Evet" + "' ";
                komut.Parameters.AddWithValue("@SiparisNo", comboBox3.Text);
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    siparisno = dr["SiparisNo"].ToString();
                    if (siparisno == comboBox3.Text)
                    {
                        izin = "var";
                        palett = dr["Palet"].ToString();
                        kesilditarihii = dr["KesildiTarihi"].ToString();
                        kesildimii = dr["KesildiMi"].ToString();
                        renkk = dr["Renk"].ToString();
                        toplamadet = dr["ToplamAdet"].ToString();
                        break;
                    }

                }

                Excel.Application excel = new Excel.Application();
                excel.Visible = true;
                object Missing = Type.Missing;
                Excel.Workbook workbook = excel.Workbooks.Open("C:\\Modeks_Dosyalar\\Etiket2-Yatay-Logosuz.xlsx");
                Excel.Worksheet sheet1 = (Excel.Worksheet)workbook.Sheets[1];

                int j2 = 2;
                int x2 = 1;
                int btr = 0;
                for (int k = 0; k < dataGridView2.Rows.Count; k++)
                {
                    int toplamAdet = Convert.ToInt32(dataGridView2.Rows[k].Cells["ToplamAdet"].Value);
                    int adet = Convert.ToInt32(dataGridView2.Rows[k].Cells["Adet"].Value);

                    for (int i = 0; i < adet; i++)
                    {
                        // satır - sütun

                        sheet1.Cells[j2 - 1, 5].Value = dataGridView2.Rows[k].Cells[0].Value.ToString(); // Sipariş Numarası
                        sheet1.Cells[j2 - 1, 4].Value = dataGridView2.Rows[k].Cells[2].Value.ToString(); // Model
                        sheet1.Cells[j2 - 1, 7].Value = toplamAdet;
                        sheet1.Cells[j2 + 1, 1].Value = dataGridView2.Rows[k].Cells[1].Value.ToString(); // Müşteri Adı
                        sheet1.Cells[j2 + 2, 1].Value = dataGridView2.Rows[k].Cells[3].Value.ToString(); // Renk
                        sheet1.Cells[j2 + 4, 6].Value = adet.ToString();
                        sheet1.Cells[j2 + 5, 4].Value = dataGridView2.Rows[k].Cells[13].Value.ToString();
                        sheet1.Cells[j2 + 7, 4].Value = dataGridView2.Rows[k].Cells[14].Value.ToString();
                        sheet1.Cells[j2 + 4, 1].Value = dataGridView2.Rows[k].Cells[15].Value.ToString();

                        sheet1.Cells[(7 + j2), 1].Value = $"{x2}/{adet}";
                        btr++;
                        if (btr != toplamAdet)
                        {
                            sheet1.Range["A1:H10"].Copy(sheet1.Range["A" + (j2 + 9) + ""]);
                        }

                        j2 += 10;

                        if (x2 == adet)
                        {
                            x2 = 1;
                        }
                        else
                        {
                            x2++;
                        }
                    }
                }


                string etiketFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Etiket");

                if (!Directory.Exists(etiketFolderPath))
                {
                    Directory.CreateDirectory(etiketFolderPath);
                }

                string excelFilePath = Path.Combine(etiketFolderPath, "Etiket" + siparisno + ".xlsx");

                try
                {
                    workbook.SaveAs(excelFilePath);

                    sheet1.PrintOutEx(Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing);
                    workbook.Close(false, Missing, Missing);
                    excel.Quit();
                    foreach (var process in Process.GetProcessesByName("EXCEL"))
                    {
                        process.Kill();
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Etiketi kaydederken bir hata oluştu, muhtemelen seçeneklerde 'hayır' butonuna bastınız, tekrar deneyiniz..");
                }

                string tarih = Convert.ToDateTime(DateTime.Now).ToString("yyyy-MM-dd HH:mm:ss");

                string sorgu3 = "UPDATE Siparişler SET Etiket=@Etiket WHERE SiparisNo=@SiparisNo AND AnaSiparişMi='" + "Evet" + "' ";
                SqlCommand komut3;
                komut3 = new SqlCommand(sorgu3, bgl.baglanti());
                komut3.Parameters.AddWithValue("@SiparisNo", comboBox3.Text);
                komut3.Parameters.AddWithValue("@Etiket", tarih);
                komut3.ExecuteNonQuery();
                label30.Text = comboBox3.Text;
                label29.Text = dataGridView1.Rows[0].Cells["Müşteri"].Value.ToString();
                Methodlar();

                textBox16.Text = "";
            }
            bgl.baglanti().Close();

        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            Form11 frm = new Form11();
            frm.Show();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Form12 frm = new Form12();
            frm.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void button23_Click(object sender, EventArgs e)
        {
            Form13 frm = new Form13();
            frm.Show();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            string TeslimTarihi = string.Empty;
            string Renk = string.Empty;
            string kayit = "";
            if (dataGridView1.Columns[e.ColumnIndex].Name == "TeslimTarihi")
            {
                TeslimTarihi = dataGridView1.Rows[e.RowIndex].Cells["TeslimTarihi"].Value?.ToString();
                Renk = dataGridView1.Rows[e.RowIndex].Cells["Renk"].Value?.ToString();
                string cnc = "Onaylandı CNCde";
                kayit = @"SELECT DISTINCT SiparisNo,Müşteri,Model,Renk,SiparişTipi,
SiparişTarihi,KesildiMi,KesildiTarihi,Onay,Palet,Etiket,ToplamM2,
ToplamAdet,CONVERT(DATETIME, TeslimTarihi, 103) AS TeslimTarihi From Siparişler 
where AnaSiparişMi='" + "Evet" + "' AND (Aşama='" + "Etiket" + "' OR (Aşama is null OR Aşama = '" + cnc + "')) AND (CONVERT(DATE, TeslimTarihi, 103) >= CONVERT(DATE, @TeslimTarihi, 103)) ORDER BY TeslimTarihi ASC";

                SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
                komut.Parameters.AddWithValue("@TeslimTarihi", TeslimTarihi);
                komut.Parameters.AddWithValue("@Renk", Renk);
                SqlDataAdapter da = new SqlDataAdapter(komut);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                onayrenklendir();

            }
            else if (dataGridView1.Columns[e.ColumnIndex].Name == "Renk")
            {
                string srg = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                string sorgu = $@"
SELECT SiparisNo,
       Müşteri,
       Model,
       Renk,
       SiparişTipi,
       SiparişTarihi,
       KesildiMi,
       KesildiTarihi,
       Onay,
       Palet,
       Etiket,
       ToplamM2,
       ToplamAdet
FROM Siparişler
WHERE Renk = '{srg}'
  AND AnaSiparişMi = 'Evet'
ORDER BY 
         Renk ASC, 
         SiparisNo ASC;
";

                SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
                DataSet ds = new DataSet();
                adap.Fill(ds, "Siparişler");
                this.dataGridView1.DataSource = ds.Tables[0];
                aynısipnugetirme();
                onayrenklendir();
            }
            else if (e.RowIndex >= 0 && dataGridView1.Columns[e.ColumnIndex].Name == "SiparisNo")
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

        private void button8_Click(object sender, EventArgs e)
        {
            liste();
            liste2();
            aynısipnugetirme();
        }
        string siparisnoacil, renkacil;
        private void TeslimTarihine3GünKalanlarıYakSöndür()
        {
            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT * FROM Siparişler WHERE CONVERT(datetime, TeslimTarihi, 104) <= DATEADD(day, 6, GETDATE()) AND Onay=@Onay AND AnaSiparişMi=@p1 AND Etiket is null ORDER BY SiparisNo ASC";
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
        string sipn;
        int satir2;
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

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
        }
        int sipnosayısı;
        int x;
        private void contextMenuStrip1_Click(object sender, EventArgs e)
        {
            sipnosayısı = 0;
            x = 0;
            Excel.Application excel = new Excel.Application();
            excel.Visible = true;
            object Missing = Type.Missing;
            //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Enes\\Desktop\\ÜretimFormu.xlsx");
            //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Modeks_Dosyalar\\ÜretimFormu.xlsx ");
            Excel.Workbook workbook = excel.Workbooks.Open("C:\\Modeks_Dosyalar\\ÜretimFormu.xlsx");

            Excel.Worksheet sheet2 = (Excel.Worksheet)workbook.Sheets[1];

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
                Excel.Range line = (Excel.Range)sheet2.Rows[11 + k];
                line.Insert();
            }

            sheet2.Cells[3, 4].Value = sipn; // siparisno yazdırma
            sheet2.Cells[4, 4].Value = dataGridView2.Rows[satir2].Cells["Müşteri"].Value.ToString(); // müşteri yazdırma
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
            foreach (var process in Process.GetProcessesByName("EXCEL"))
            {
                process.Kill();
            }
            bgl.baglanti().Close();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            rengegöresırala();
            aynısipnugetirme();
        }

        private void button24_Click(object sender, EventArgs e)
        {
            Form14 frm = new Form14();
            frm.ShowDialog();
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            Form7 frm = new Form7();
            this.Hide();
            frm.yetki = yetki;
            frm.Show();
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void Form7_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Hide();
            Form1 form1 = Application.OpenForms["Form1"] as Form1;
            if (form1 != null)
            {
                form1.Show();
            }
        }
        private void bugünegöresırala()
        {
            DateTime bitir = DateTime.Now;
            DateTime basla = DateTime.Now;
            dateTimePicker1.Value = basla;
            label27.Text = basla.ToString("yyyy - MM - dd");
            label28.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");
            string sorgu = "SELECT DISTINCT SiparisNo,Müşteri,Model,Renk,SiparişTipi,SiparişTarihi,KesildiMi,KesildiTarihi,Onay,Palet,Etiket,ToplamM2,ToplamAdet From Siparişler where SiparişTarihi between '" + label27.Text + "' AND '" + label28.Text + "' AND  AnaSiparişMi='Evet' AND Aşama='Etiket' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
        }
        private void haftayagöresırala()
        {
            DateTime bitir = DateTime.Now;
            DateTime basla = bitir.AddDays(-7);
            dateTimePicker1.Value = basla;
            label27.Text = basla.ToString("yyyy - MM - dd HH:mm:ss");
            label28.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");
            string sorgu = "SELECT DISTINCT SiparisNo,Müşteri,Model,Renk,SiparişTipi,SiparişTarihi,KesildiMi,KesildiTarihi,Onay,Palet,Etiket,ToplamM2,ToplamAdet From Siparişler where SiparişTarihi between '" + label27.Text + "' AND '" + label28.Text + "' AND  AnaSiparişMi='Evet' AND Aşama='Etiket' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
        }
        private void ayagöresırala()
        {
            DateTime bitir = DateTime.Now;
            DateTime basla = bitir.AddMonths(-1);
            dateTimePicker1.Value = basla;
            label27.Text = basla.ToString("yyyy - MM - dd HH:mm:ss");
            label28.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");
            string sorgu = "SELECT DISTINCT SiparisNo,Müşteri,Model,Renk,SiparişTipi,SiparişTarihi,KesildiMi,KesildiTarihi,Onay,Palet,Etiket,ToplamM2,ToplamAdet From Siparişler where SiparişTarihi between '" + label27.Text + "' AND '" + label28.Text + "' AND  AnaSiparişMi='Evet' AND Aşama='Etiket' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];

        }
        private void yılagöresırala()
        {
            DateTime bitir = DateTime.Now;
            DateTime basla = bitir.AddYears(-1);
            dateTimePicker1.Value = basla;
            label27.Text = basla.ToString("yyyy - MM - dd HH:mm:ss");
            label28.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");
            string sorgu = "SELECT DISTINCT SiparisNo,Müşteri,Model,Renk,SiparişTipi,SiparişTarihi,KesildiMi,KesildiTarihi,Onay,Palet,Etiket,ToplamM2,ToplamAdet From Siparişler where SiparişTarihi between '" + label27.Text + "' AND '" + label28.Text + "' AND  AnaSiparişMi='Evet' AND Aşama='Etiket' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];

        }
        private void button1_Click(object sender, EventArgs e)
        {
            bugünegöresırala();
            aynısipnugetirme();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            haftayagöresırala();
            aynısipnugetirme();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ayagöresırala();
            aynısipnugetirme();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            yılagöresırala();
            aynısipnugetirme();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTime bitir = dateTimePicker1.Value;
            DateTime basla = dateTimePicker2.Value;
            label27.Text = basla.ToString("yyyy - MM - dd");
            label28.Text = bitir.ToString("yyyy - MM - dd");
            string sorgu = "SELECT DISTINCT SiparisNo,Müşteri,Model,Renk,SiparişTipi,SiparişTarihi,KesildiMi,KesildiTarihi,Onay,Palet,Etiket,ToplamM2,ToplamAdet From Siparişler where SiparişTarihi between '" + label27.Text + "' AND '" + label28.Text + "' AND  AnaSiparişMi='Evet' AND Aşama='Etiket' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            DateTime bitir = dateTimePicker1.Value;
            DateTime basla = dateTimePicker2.Value;
            label27.Text = basla.ToString("yyyy - MM - dd");
            label28.Text = bitir.ToString("yyyy - MM - dd");
            string sorgu = "SELECT DISTINCT SiparisNo,Müşteri,Model,Renk,SiparişTipi,SiparişTarihi,KesildiMi,KesildiTarihi,Onay,Palet,Etiket,ToplamM2,ToplamAdet From Siparişler where SiparişTarihi between '" + label27.Text + "' AND '" + label28.Text + "' AND  AnaSiparişMi='Evet' AND Aşama='Etiket' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
        }

        private void button26_Click(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            liste();
            liste2();
            aynısipnugetirme();
        }

        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            aynısipnugetirme();
            onayrenklendir();
        }
        double pressyollanansayısı = 0;
        double pressyollananmetresi = 0;
        double pressyollananadeti = 0;

        private void PressYollanan()
        {
            pressyollanansayısı = 0;
            pressyollananmetresi = 0;
            pressyollananadeti = 0;
            SqlCommand cmd = new SqlCommand("select count(distinct SiparisNo) from Siparişler Where Etiket is not null and MembranPressTarihi is null and Aşama=@Aşama", bgl.baglanti());
            cmd.Parameters.AddWithValue("@Aşama", "Palet");
            pressyollanansayısı = Convert.ToInt32(cmd.ExecuteScalar());

            SqlCommand komut = new SqlCommand();
            komut.CommandText = "select * from Siparişler Where Etiket is not null and MembranPressTarihi is null and AnaSiparişMi=@AnaSiparişMi and Aşama=@Aşama";
            komut.Parameters.AddWithValue("@AnaSiparişMi", "Evet");
            komut.Parameters.AddWithValue("@Aşama", "Palet");
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                pressyollananmetresi += Convert.ToDouble(dr["M2"]);
                pressyollananadeti += Convert.ToDouble(dr["Adet"]);
            }
            textBox32.Text = pressyollanansayısı.ToString();
            textBox33.Text = pressyollananmetresi.ToString();
            textBox41.Text = pressyollananadeti.ToString();
        }

        private void button27_Click(object sender, EventArgs e)
        {
            Form18 frm = new Form18();
            frm.yetki = yetki;
            frm.ShowDialog();
        }

        private void textBox49_TextChanged(object sender, EventArgs e)
        {
            if (textBox49.Text == "")
            {
                Methodlar();
            }
            else
            {
                try
                {
                    string kayit = "SELECT DISTINCT SiparisNo, Müşteri, Model, Renk, SiparişTipi, SiparişTarihi, KesildiMi, KesildiTarihi, Onay, Palet, Etiket, ToplamM2, ToplamAdet FROM Siparişler WHERE AnaSiparişMi = @p1 AND Aşama = @Aşama AND Palet LIKE '' + @searchText + '' ORDER BY SiparisNo DESC";

                    SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
                    komut.Parameters.AddWithValue("@p1", "Evet");
                    komut.Parameters.AddWithValue("@Aşama", "Etiket");
                    komut.Parameters.AddWithValue("@searchText", textBox49.Text);

                    SqlDataAdapter da = new SqlDataAdapter(komut);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;

                    string kayit2 = "SELECT SiparisNo, Müşteri, Model, Renk, SiparişTipi, SiparişTarihi, KesildiMi, KesildiTarihi, Onay, Palet, Etiket, ToplamM2, ToplamAdet, Boy, En, Özellik, Telefon, TeslimTarihi, Adres, M2, Adet FROM Siparişler WHERE Aşama = @Aşama AND Palet LIKE '' + @searchText + '' ORDER BY SiparisNo ASC, Boy DESC";

                    SqlCommand komut2 = new SqlCommand(kayit2, bgl.baglanti());
                    komut2.Parameters.AddWithValue("@Aşama", "Etiket");
                    komut2.Parameters.AddWithValue("@searchText", textBox49.Text);

                    SqlDataAdapter da2 = new SqlDataAdapter(komut2);
                    DataTable dt2 = new DataTable();
                    da2.Fill(dt2);
                    dataGridView2.DataSource = dt2;

                    aynısipnugetirme();
                }
                catch (Exception)
                {
                    Methodlar();
                }
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            string kayit = "SELECT DISTINCT SiparisNo,Müşteri,Model,Renk,SiparişTipi,SiparişTarihi,KesildiMi,KesildiTarihi,Onay,Palet,Etiket,ToplamM2,ToplamAdet,TeslimTarihi From Siparişler where  AnaSiparişMi=@p1 AND Aşama=@Aşama ORDER BY TeslimTarihi ASC";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            komut.Parameters.AddWithValue("@p1", "Evet");
            komut.Parameters.AddWithValue("@Aşama", "Etiket");
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            üçgüniçindeki();
            bgl.baglanti().Close();


            aynısipnugetirme();
        }
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
        private void button25_Click(object sender, EventArgs e)
        {
            MEMBRANDANGERIAL();
        }

        private void groupBox2_Enter(object sender, EventArgs e)
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
