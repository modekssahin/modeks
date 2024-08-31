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
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using FirebirdSql.Data.FirebirdClient;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.SqlServer.Management.HadrModel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using System.Diagnostics;
using System.Windows.Media.Animation;

namespace Modeks
{
    public partial class Form10 : Form
    {
        public Form10()
        {
            InitializeComponent();
        }
        sqlsinif bgl = new sqlsinif();
        public string yetki;
        public string kullaniciadi;
        public string hangiformdan;
        public string hangiformdanModeksEski;
        public string SiparisNoSuEski;
        public string MusteriEski;
        public string SiparisTarihi;
        SqlCommand cmd;
        public string siparişno = "";
        public bool vbkontrol;
        string id;
        double boy;
        double en;
        double adet;
        double m2;
        double toplamM2;
        double toplamKapakAdet;
        double toplamTasarımUcreti;
        double m2kapaksayısı;
        double kapaktoplam;
        double fiyat1;
        double toplamfiyat = 0;
        double geneltoplam = 0;
        double iskonto = 0;
        double acil = 0;
        double m2kapakfarkı = 0;
        double kargo = 0;
        double dds = 0;
        int kayitSayisi;
        int bölmesayısı = 0;
        int bölmekayitsayisi = 0;
        string siparisbölme;

        string message;

        double istatistiktoplamadet;
        double istatistiktoplamm2;
        private string selectedText = "";
        public void müsteri_cek()
        {
            try
            {
                // Önce ListBox'ı temizle
                comboBox1.Items.Clear();
                var connectionString = @"User ID=SYSDBA;Password=masterkey;Database=78.108.246.74:C:\Program Files (x86)\BSR\VERESIYEDATA.FDB ;Charset=WIN1254;";
                FbConnection fbcnn = new FbConnection(connectionString);
                fbcnn.Open();

                string sql = "SELECT * FROM MUSTERILER WHERE GIZLE='0'";
                FbCommand command = new FbCommand(sql, fbcnn);
                FbDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    comboBox1.Items.Add(reader["ADI_SOYADI"].ToString());
                }
                label66.Text = "KURULDU";
                label66.ForeColor = Color.Green;
            }
            catch (Exception)
            {
                label66.Text = "KURULAMADI";
                label66.ForeColor = Color.Red;
            }
        }

        int BSR_ID;
        double ODEME;
        double SATIS_TUTARI;
        double BAKİYE;
        public void müsteri_bilgileri_getir()
        {
            try
            {
                var connectionString = @"User ID=SYSDBA;Password=masterkey;Database=78.108.246.74:C:\Program Files (x86)\BSR\VERESIYEDATA.FDB ;Charset=WIN1254;";
                FbConnection fbcnn = new FbConnection(connectionString);//bağlan
                string sql = "select * from MUSTERILER where ADI_SOYADI='" + comboBox1.Text + "'";
                fbcnn.Open();
                FbCommand command = new FbCommand(sql, fbcnn);//komut
                FbDataReader reader = command.ExecuteReader();//Reader yaparak Datareaderin içerisine aktar.
                StringBuilder sb = new StringBuilder();
                while (reader.Read())
                {
                    //textBox3.Text = reader["ADRES"].ToString();
                    Byte[] byteData = (Byte[])(reader["ADRES"]);
                    String adres = System.Text.Encoding.UTF8.GetString(byteData);
                    textBox3.Text = adres;
                    textBox4.Text = reader["TELEFON"].ToString();
                    textBox5.Text = reader["EPOSTA"].ToString();
                    BSR_ID = Convert.ToInt32(reader["ID"].ToString());
                    //label62.Text = Convert.ToInt32(reader["ODEMELER"].ToString())  reader["SATIS"].ToString() + "TL";
                    ODEME = Convert.ToDouble(reader["ODEMELER"].ToString());
                    SATIS_TUTARI = Convert.ToDouble(reader["SATIS"].ToString());
                    double limit = Convert.ToDouble(reader["LIMIT"].ToString());

                    //BAKİYE = Convert.ToDouble(reader["SATIS"].ToString());
                    BAKİYE = SATIS_TUTARI - ODEME;
                    label62.Text = Convert.ToDecimal(BAKİYE).ToString("N2") + " TL";
                    label78.Text = Convert.ToDecimal(limit).ToString("N2") + " TL";
                    if (limit != 0 && limit * 2 < (SATIS_TUTARI - ODEME))
                    {
                        MessageBox.Show("Dikkat! Bu müşteri bakiye limitinin 2 katını aşmıştır. Müşteriye sipariş oluşturamazsınız.\n\nBakiye Limiti: " + limit);
                        Form3 frm = new Form3();
                        this.Hide();
                        frm.ShowDialog();
                    }
                }
            }
            catch (Exception)
            {
                label66.Text = "KURULAMADI";
                label66.ForeColor = Color.Red;
            }

        }
        int Satıs_ID = 100000;
        private void satıs_ıd_getir()
        {
            try
            {
                var connectionString = @"User ID=SYSDBA;Password=masterkey;Database=localhost:C:\Program Files (x86)\BSR\VERESIYEDATA.FDB ;Charset=WIN1254;";
                FbConnection fbcnn = new FbConnection(connectionString);//bağlan
                string sql = "select * from SATISLAR";
                fbcnn.Open();
                FbCommand command = new FbCommand(sql, fbcnn);//komut
                FbDataReader reader = command.ExecuteReader();//Reader yaparak Datareaderin içerisine aktar.
                StringBuilder sb = new StringBuilder();
                while (reader.Read())
                {
                    if (Satıs_ID < Convert.ToInt32(reader["ID"].ToString()))
                    {
                        Satıs_ID = Convert.ToInt32(reader["ID"]);
                    }

                }
            }
            catch (Exception)
            {

                label66.Text = "KURULAMADI";
                label66.ForeColor = Color.Red;
            }

        }
        int BSR_KULLANICI_ID;
        private void bsr_kullanici_getir()
        {
            try
            {
                var connectionString = @"User ID=SYSDBA;Password=masterkey;Database=78.108.246.74:C:\Program Files (x86)\BSR\VERESIYEDATA.FDB ;Charset=WIN1254;";
                FbConnection fbcnn = new FbConnection(connectionString);//bağlan
                string sql = "select * from KULLANICILAR where KULLANICIADI='" + kullaniciadi + "'";
                fbcnn.Open();
                FbCommand command = new FbCommand(sql, fbcnn);//komut
                FbDataReader reader = command.ExecuteReader();//Reader yaparak Datareaderin içerisine aktar.
                StringBuilder sb = new StringBuilder();
                while (reader.Read())
                {
                    BSR_KULLANICI_ID = Convert.ToInt32(reader["ID"].ToString());
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        Double BSR_BIRIMFIYAT;
        private void BSR_Kaydet()
        {
            satıs_ıd_getir();
            bsr_kullanici_getir();
            try
            {
                var connectionString = @"User ID=SYSDBA;Password=masterkey;Database=78.108.246.74:C:\Program Files (x86)\BSR\VERESIYEDATA.FDB ;Charset=WIN1254;";
                FbConnection fbcnn = new FbConnection(connectionString);//bağlan
                fbcnn.Open();
                string kayit = "insert into SATISLAR(TUR_ID,MUSTERI_ID,ID,ALINANLAR,TARIH,MIKTAR,ADET,TOPLAM,TUR,BIRIM,ACIKLAMA,KULLANICI_ID) values (@TUR_ID,@MUSTERI_ID,@ID,@ALINANLAR,@TARIH,@MIKTAR,@ADET,@TOPLAM,@TUR,@BIRIM,@ACIKLAMA,@KULLANICI_ID)";
                FbCommand komut = new FbCommand(kayit, fbcnn);
                komut.Parameters.AddWithValue("@TUR_ID", 1);
                komut.Parameters.AddWithValue("@MUSTERI_ID", BSR_ID);
                komut.Parameters.AddWithValue("@ID", (Satıs_ID + 1));
                komut.Parameters.AddWithValue("@ALINANLAR", "PVC");
                komut.Parameters.AddWithValue("@TARIH", DateTime.Now.ToString());
                BSR_BIRIMFIYAT = (kapaktoplam / toplamM2);
                komut.Parameters.AddWithValue("@MIKTAR", BSR_BIRIMFIYAT);
                komut.Parameters.AddWithValue("@ADET", toplamM2);
                komut.Parameters.AddWithValue("@TOPLAM", geneltoplam);
                komut.Parameters.AddWithValue("@TUR", "SATIS");
                komut.Parameters.AddWithValue("@BIRIM", "ADET");
                komut.Parameters.AddWithValue("@ACIKLAMA", textBox1.Text +" -  Ekleyen Kullanıcı: "+ kullaniciadi);
                komut.Parameters.AddWithValue("@KULLANICI_ID", BSR_KULLANICI_ID);
                komut.ExecuteNonQuery();
                fbcnn.Close();
                //MessageBox.Show("Bsr'ye satış eklenmiştir ve onaylanmıştır.");
            }
            catch (Exception)
            {
                var connectionString = @"User ID=SYSDBA;Password=masterkey;Database=78.108.246.74:C:\Program Files (x86)\BSR\VERESIYEDATA.FDB ;Charset=WIN1254;";
                FbConnection fbcnn = new FbConnection(connectionString);//bağlan
                fbcnn.Open();
                string kayit = "insert into SATISLAR(TUR_ID,MUSTERI_ID,ID,ALINANLAR,TARIH,MIKTAR,ADET,TOPLAM,TUR,BIRIM,ACIKLAMA,KULLANICI_ID) values (@TUR_ID,@MUSTERI_ID,@ID,@ALINANLAR,@TARIH,@MIKTAR,@ADET,@TOPLAM,@TUR,@BIRIM,@ACIKLAMA,@KULLANICI_ID)";
                FbCommand komut = new FbCommand(kayit, fbcnn);
                komut.Parameters.AddWithValue("@TUR_ID", 1);
                komut.Parameters.AddWithValue("@MUSTERI_ID", BSR_ID);
                komut.Parameters.AddWithValue("@ID", (Satıs_ID + 2));
                komut.Parameters.AddWithValue("@ALINANLAR", "PVC");
                komut.Parameters.AddWithValue("@TARIH", DateTime.Now.ToString());
                BSR_BIRIMFIYAT = (kapaktoplam / toplamM2);
                komut.Parameters.AddWithValue("@MIKTAR", BSR_BIRIMFIYAT);
                komut.Parameters.AddWithValue("@ADET", toplamM2);
                komut.Parameters.AddWithValue("@TOPLAM", geneltoplam);
                komut.Parameters.AddWithValue("@TUR", "SATIS");
                komut.Parameters.AddWithValue("@BIRIM", "ADET");
                komut.Parameters.AddWithValue("@ACIKLAMA", textBox1.Text);
                komut.Parameters.AddWithValue("@KULLANICI_ID", BSR_KULLANICI_ID);
                komut.ExecuteNonQuery();
                fbcnn.Close();
            }
        }

        private void liste()
        {
            string kayit = "SELECT id as 'ID', Model,Renk,TasarımÜcreti,Özellik,Boy,En,Adet,M2,M2Fiyat,Fiyat2,BID,Nott as 'Not' From Siparişler where SiparisNo=@SiparisNo";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            komut.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
            SqlDataAdapter da = new SqlDataAdapter(komut);
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
        }
        private void modelbilgilerigetir()
        {
            try
            {
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Modeller";
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    comboBox4.Items.Add(dr["ModelAdı"]);
                }
            }
            catch (Exception)
            {

                throw;
            }

        }
        private void renkbilgilerigetir()
        {
            try
            {
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Renkler";
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    comboBox6.Items.Add(dr["RenkAdı"]);
                }
            }
            catch (Exception)
            {

                throw;
            }

        }
        private void özellikbilgilerigetir()
        {
            try
            {
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Özellik";
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    comboBox7.Items.Add(dr["ÖzellikAdı"]);
                }
            }
            catch (Exception)
            {

                throw;
            }

        }
        private void fiyatbilgisigetir()
        {
            try
            {
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Modeller where ModelAdı=@ModelAdı";
                komut.Parameters.AddWithValue("@ModelAdı", comboBox4.Text);
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    fiyat1 = Convert.ToDouble(dr["MaTFiyat"]);
                    textBox15.Text = fiyat1.ToString();
                }
            }
            catch (Exception)
            {
                //MessageBox.Show("Bu modelin matı yoktur.", "Bilgilendirme", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        private void hgfiyatgetir()
        {
            try
            {
                if (comboBox6.Text.StartsWith("HG"))
                {
                    SqlCommand komut = new SqlCommand();
                    komut.CommandText = "SELECT *FROM Modeller where ModelAdı=@ModelAdı";
                    komut.Parameters.AddWithValue("@ModelAdı", comboBox4.Text);
                    komut.Connection = bgl.baglanti();
                    komut.CommandType = CommandType.Text;

                    SqlDataReader dr;
                    dr = komut.ExecuteReader();
                    while (dr.Read())
                    {
                        fiyat1 = Convert.ToDouble(dr["HgFiyat"]);
                        textBox15.Text = fiyat1.ToString();
                    }
                }
            }
            catch (Exception)
            {
                //MessageBox.Show("Bu modelin parlağı yoktur.", "Bilgilendirme", MessageBoxButtons.OK, MessageBoxIcon.Error);
                comboBox4.SelectedIndex = -1;
                comboBox6.SelectedIndex = -1;
            }

        }
        private void RSMfiyatgetir()
        {
            try
            {
                if (comboBox6.Text.StartsWith("RSM"))
                {
                    SqlCommand komut = new SqlCommand();
                    komut.CommandText = "SELECT *FROM Modeller where ModelAdı=@ModelAdı";
                    komut.Parameters.AddWithValue("@ModelAdı", comboBox4.Text);
                    komut.Connection = bgl.baglanti();
                    komut.CommandType = CommandType.Text;

                    SqlDataReader dr;
                    dr = komut.ExecuteReader();
                    while (dr.Read())
                    {
                        fiyat1 = Convert.ToDouble(dr["RSM"]);
                        textBox15.Text = fiyat1.ToString();
                    }
                }
            }
            catch (Exception)
            {
                //MessageBox.Show("Bu modelin parlağı yoktur.", "Bilgilendirme", MessageBoxButtons.OK, MessageBoxIcon.Error);
                comboBox4.SelectedIndex = -1;
                comboBox6.SelectedIndex = -1;
            }

        }
        private void SOFTfiyatgetir()
        {
            try
            {
                if (comboBox6.Text.StartsWith("*"))
                {
                    SqlCommand komut = new SqlCommand();
                    komut.CommandText = "SELECT *FROM Modeller where ModelAdı=@ModelAdı";
                    komut.Parameters.AddWithValue("@ModelAdı", comboBox4.Text);
                    komut.Connection = bgl.baglanti();
                    komut.CommandType = CommandType.Text;

                    SqlDataReader dr;
                    dr = komut.ExecuteReader();
                    while (dr.Read())
                    {
                        fiyat1 = Convert.ToDouble(dr["SOFT"]);
                        textBox15.Text = fiyat1.ToString();
                    }
                }
            }
            catch (Exception)
            {
                //MessageBox.Show("Bu modelin parlağı yoktur.", "Bilgilendirme", MessageBoxButtons.OK, MessageBoxIcon.Error);
                comboBox4.SelectedIndex = -1;
                comboBox6.SelectedIndex = -1;
            }

        }
        private void ULTRASOFTfiyatgetir()
        {
            try
            {
                if (comboBox6.Text.StartsWith("**"))
                {
                    SqlCommand komut = new SqlCommand();
                    komut.CommandText = "SELECT *FROM Modeller where ModelAdı=@ModelAdı";
                    komut.Parameters.AddWithValue("@ModelAdı", comboBox4.Text);
                    komut.Connection = bgl.baglanti();
                    komut.CommandType = CommandType.Text;

                    SqlDataReader dr;
                    dr = komut.ExecuteReader();
                    while (dr.Read())
                    {
                        fiyat1 = Convert.ToDouble(dr["ULTRASOFT"]);
                        textBox15.Text = fiyat1.ToString();
                    }
                }
            }
            catch (Exception)
            {
                //MessageBox.Show("Bu modelin parlağı yoktur.", "Bilgilendirme", MessageBoxButtons.OK, MessageBoxIcon.Error);
                comboBox4.SelectedIndex = -1;
                comboBox6.SelectedIndex = -1;
            }

        }
        private void resimgetir()
        {
            try
            {
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Modeller where ModelAdı=@ModelAdı";
                komut.Parameters.AddWithValue("@ModelAdı", comboBox4.Text);
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    pictureBox2.ImageLocation = dr["resimyolu"].ToString();

                }

            }
            catch (Exception)
            {

                throw;
            }

        }
        private void resimrenkgetir()
        {
            try
            {
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Renkler where RenkAdı=@RenkAdı";
                komut.Parameters.AddWithValue("@RenkAdı", comboBox6.Text);
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    pictureBox3.ImageLocation = dr["renkyolu"].ToString();

                }

            }
            catch (Exception)
            {

                throw;
            }

        }
        private void iskontogetir()
        {
            try
            {
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Siparişler where SiparisNo=@p1 AND AnaSiparişMi=@p2";
                komut.Parameters.AddWithValue("@p1", textBox1.Text);
                komut.Parameters.AddWithValue("@p2", "Evet");

                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    if (dr["İskonto"].ToString() == "") label42.Text = "0";
                    else label42.Text =  dr["İskonto"].ToString();

                    if (dr["İskontoOrani"].ToString() == "") textBox16.Text = "0";
                    else textBox16.Text = dr["İskontoOrani"].ToString();
                }
            }
            catch (Exception)
            {

                throw;
            }

        }
        double özellikfiyat;
        private void özellikfiyatekle()
        {
            try
            {
                özellikfiyat = 0;
                fiyatbilgisigetir();
                hgfiyatgetir();
                RSMfiyatgetir();
                SOFTfiyatgetir();
                ULTRASOFTfiyatgetir();

                if (comboBox7.Text != "CEKMECE BUTUN" && comboBox7.Text != "CIZIM")
                {
                    SqlCommand komut = new SqlCommand();
                    komut.CommandText = "SELECT *FROM Modeller where ModelAdı=@ModelAdı";
                    komut.Parameters.AddWithValue("@ModelAdı", comboBox4.Text);
                    komut.Connection = bgl.baglanti();
                    komut.CommandType = CommandType.Text;

                    SqlDataReader dr;
                    dr = komut.ExecuteReader();
                    while (dr.Read())
                    {
                        özellikfiyat = (Convert.ToDouble(textBox15.Text) * Convert.ToDouble(dr[comboBox7.Text].ToString())) / 100;
                        fiyat1 = Convert.ToDouble(textBox15.Text) + özellikfiyat;
                        textBox15.Text = fiyat1.ToString();
                    }
                }
            }
            catch (Exception)
            {

            }

        }
        private void formül()
        {
            m2 = (en * boy * adet) / 10000;
            textBox12.Text = m2.ToString("0.##");
        }
        private void kayıtsayısı()
        {
            try
            {
                SqlCommand cmd = new SqlCommand("select top 1 SiparisNo from Siparişler ORDER BY SiparisNo DESC;", bgl.baglanti());
                kayitSayisi = Convert.ToInt32(cmd.ExecuteScalar());
            }
            catch (Exception)
            {
                SqlCommand cmd = new SqlCommand("SELECT top 1 LEFT(SiparisNo,3)  FROM Siparişler WHERE SiparisNo LIKE '%-%'  ORDER BY SiparisNo DESC;", bgl.baglanti());
                kayitSayisi = Convert.ToInt32(cmd.ExecuteScalar());
            }

        }
        private void İstatistik()
        {
            try
            {
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM İstatistikler where RenkAdı=@RenkAdı";
                komut.Parameters.AddWithValue("@RenkAdı", comboBox6.Text);
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())

                {
                    istatistiktoplamadet = Convert.ToDouble(dr["SatışAdedi"]);
                    istatistiktoplamm2 = Convert.ToDouble(dr["ToplamM2"]);
                }

                istatistiktoplamadet += 1;
                istatistiktoplamm2 += Convert.ToDouble(textBox12.Text);
                string sorgu2 = "UPDATE İstatistikler SET ToplamM2=@ToplamM2,SatışAdedi=@SatışAdedi WHERE RenkAdı=@RenkAdı";
                SqlCommand komut2;
                komut2 = new SqlCommand(sorgu2, bgl.baglanti());
                komut2.Parameters.AddWithValue("@RenkAdı", comboBox6.Text);
                komut2.Parameters.AddWithValue("@ToplamM2", istatistiktoplamm2.ToString("0.##"));
                komut2.Parameters.AddWithValue("@SatışAdedi", istatistiktoplamadet.ToString("0.##"));
                komut2.ExecuteNonQuery();
            }
            catch (Exception)
            {

                throw;
            }

        }
        private void kaydet()
        {
            if (dataGridView1.Rows.Count == 1)
            {
                if (textBox14.Text == "")
                {
                    textBox14.Text = "0";
                }
                if (comboBox4.Text != "" && comboBox6.Text != "" && comboBox2.Text != "" && comboBox3.Text != "" && comboBox1.Text != "")
                {
                    if (textBox2.Text != "0" && textBox6.Text != "0" && textBox7.Text != "0" && textBox2.Text != "" && textBox6.Text != "" && textBox7.Text != "")
                    {
                        // 2 li
                        string sorgu = "SELECT COUNT(*) FROM Siparişler WHERE SiparisNo =@p1";
                        SqlCommand komut;
                        komut = new SqlCommand(sorgu, bgl.baglanti());
                        komut.Parameters.AddWithValue("@p1", textBox1.Text);
                        int count = (int)komut.ExecuteScalar();
                        if (count > 0)
                        {
                            textBox1.Text = Convert.ToString(Convert.ToDouble(textBox1.Text) + 1);
                        }
                        else
                        {

                        }

                        //3 lü 
                        string sorgu2 = "SELECT COUNT(*) FROM Siparişler WHERE SiparisNo =@p1";
                        SqlCommand komut2;
                        komut2 = new SqlCommand(sorgu2, bgl.baglanti());
                        komut2.Parameters.AddWithValue("@p1", textBox1.Text);
                        int count2 = (int)komut2.ExecuteScalar();
                        if (count2 > 0)
                        {
                            textBox1.Text = Convert.ToString(Convert.ToDouble(textBox1.Text) + 1);
                        }
                        else
                        {

                        }
                        try
                        {
                            string tarih = Convert.ToDateTime(dateTimePicker1.Value).ToString("yyyy-MM-dd HH:mm:ss");
                            string teslimtarihi = Convert.ToDateTime(textBox10.Text).ToString("dd.MM.yyyy HH:mm:ss");

                            SqlCommand cmd2 = new SqlCommand();
                            cmd2.Connection = bgl.baglanti();

                            cmd2.CommandText = "INSERT INTO Siparişler (SiparisNo, Onay, Müşteri, SiparişTipi, SevkTürü, BID, Adres, Telefon, SiparişTarihi, TeslimTarihi, Nott, Model, Renk, Fiyat, TasarımÜcreti, Özellik, BaskıYönü, Boy, En, Adet, M2, M2Fiyat, Fiyat2, AnaSiparişMi, İskonto, EkleyenKullanici) VALUES " +
                                               "(@SiparisNo, @Onay, @Müşteri, @SiparişTipi, @SevkTürü, @BID, @Adres, @Telefon, @SiparişTarihi, @TeslimTarihi, @Nott, @Model, @Renk, @Fiyat, @TasarımÜcreti, @Özellik, @BaskıYönü, @Boy, @En, @Adet, @M2, @M2Fiyat, @Fiyat2, @AnaSiparişMi, @İskonto, @EkleyenKullanici)";

                            cmd2.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
                            cmd2.Parameters.AddWithValue("@Onay", "Onay Bekliyor");
                            cmd2.Parameters.AddWithValue("@Müşteri", comboBox1.Text);
                            cmd2.Parameters.AddWithValue("@SiparişTipi", comboBox2.Text);
                            cmd2.Parameters.AddWithValue("@SevkTürü", comboBox3.Text);
                            cmd2.Parameters.AddWithValue("@BID", textBox8.Text);
                            cmd2.Parameters.AddWithValue("@Adres", textBox3.Text);
                            cmd2.Parameters.AddWithValue("@Telefon", textBox4.Text);
                            cmd2.Parameters.AddWithValue("@SiparişTarihi", tarih);
                            cmd2.Parameters.AddWithValue("@TeslimTarihi", teslimtarihi);
                            cmd2.Parameters.AddWithValue("@Nott", textBox11.Text);
                            cmd2.Parameters.AddWithValue("@Model", comboBox4.Text);
                            cmd2.Parameters.AddWithValue("@Renk", comboBox6.Text);
                            cmd2.Parameters.AddWithValue("@Fiyat", textBox15.Text);
                            cmd2.Parameters.AddWithValue("@TasarımÜcreti", textBox14.Text);
                            cmd2.Parameters.AddWithValue("@Özellik", comboBox7.Text);
                            cmd2.Parameters.AddWithValue("@BaskıYönü", comboBox5.Text);
                            cmd2.Parameters.AddWithValue("@Boy", textBox2.Text);
                            cmd2.Parameters.AddWithValue("@En", textBox6.Text);
                            cmd2.Parameters.AddWithValue("@Adet", textBox7.Text);
                            cmd2.Parameters.AddWithValue("@M2", textBox12.Text);
                            cmd2.Parameters.AddWithValue("@M2Fiyat", textBox15.Text);
                            cmd2.Parameters.AddWithValue("@Fiyat2", textBox17.Text);
                            cmd2.Parameters.AddWithValue("@AnaSiparişMi", "Evet");
                            cmd2.Parameters.AddWithValue("@İskonto", textBox16.Text);
                            cmd2.Parameters.AddWithValue("@EkleyenKullanici", kullaniciadi);


                            cmd2.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }

                        İstatistik();
                        liste();
                        kayıtsayısı();
                        M2Toplat();
                        KapakAdetToplat();
                        TasarımUcretıToplat();
                        KapakToplam();
                        M2KapakSayısı();
                        if (comboBox2.Text == "Acil")
                        {
                            AcilHesaplama();
                        }
                        M2KapakFarkıHesaplama();
                        KargoHesaplama();
                        Toplam();
                        AraToplam();
                        DDSHesaplama();
                        GenelToplam();
                        StokGetir();
                        TutkalStok();
                        MDFStok();
                        temizle();
                    }
                    else
                    {
                        MessageBox.Show("Ekleme yapabilmek için en, boy, adet bilgilerini yazın.", "Mesaj", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }

                }
                else
                {
                    MessageBox.Show("Ekleme yapabilmek için Müşteri, Sipariş Şekli, Sevk Türü,  Model, Renk bilgilerini doldurun.", "Mesaj", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            else
            {
                if (textBox14.Text == "")
                {
                    textBox14.Text = "0";
                }
                if (textBox2.Text != "0" && textBox6.Text != "0" && textBox7.Text != "0")
                {

                    string tarih = Convert.ToDateTime(DateTime.Now).ToString("yyyy-MM-dd HH:mm:ss");
                    string teslimtarihi = Convert.ToDateTime(textBox10.Text).ToString("dd.MM.yyyy HH:mm:ss");
                    SqlCommand cmd4 = new SqlCommand();
                    cmd4.Connection = bgl.baglanti();

                    cmd4.CommandText = "INSERT INTO Siparişler (SiparisNo, Onay, Müşteri, SiparişTipi, SevkTürü, BID, Adres, Telefon, SiparişTarihi, TeslimTarihi, Nott, Model, Renk, Fiyat, TasarımÜcreti, Özellik, BaskıYönü, Boy, En, Adet, M2, M2Fiyat, Fiyat2, AnaSiparişMi, İskonto, EkleyenKullanici) VALUES " +
                                       "(@SiparisNo, @Onay, @Müşteri, @SiparişTipi, @SevkTürü, @BID, @Adres, @Telefon, @SiparişTarihi, @TeslimTarihi, @Nott, @Model, @Renk, @Fiyat, @TasarımÜcreti, @Özellik, @BaskıYönü, @Boy, @En, @Adet, @M2, @M2Fiyat, @Fiyat2, @AnaSiparişMi, @İskonto, @EkleyenKullanici)";

                    cmd4.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
                    cmd4.Parameters.AddWithValue("@Onay", "Onay Bekliyor");
                    cmd4.Parameters.AddWithValue("@Müşteri", comboBox1.Text);
                    cmd4.Parameters.AddWithValue("@SiparişTipi", comboBox2.Text);
                    cmd4.Parameters.AddWithValue("@SevkTürü", comboBox3.Text);
                    cmd4.Parameters.AddWithValue("@BID", textBox8.Text);
                    cmd4.Parameters.AddWithValue("@Adres", textBox3.Text);
                    cmd4.Parameters.AddWithValue("@Telefon", textBox4.Text);
                    cmd4.Parameters.AddWithValue("@SiparişTarihi", tarih);
                    cmd4.Parameters.AddWithValue("@TeslimTarihi", teslimtarihi);
                    cmd4.Parameters.AddWithValue("@Nott", textBox11.Text);
                    cmd4.Parameters.AddWithValue("@Model", comboBox4.Text);
                    cmd4.Parameters.AddWithValue("@Renk", comboBox6.Text);
                    cmd4.Parameters.AddWithValue("@Fiyat", textBox15.Text);
                    cmd4.Parameters.AddWithValue("@TasarımÜcreti", textBox14.Text);
                    cmd4.Parameters.AddWithValue("@Özellik", comboBox7.Text);
                    cmd4.Parameters.AddWithValue("@BaskıYönü", comboBox5.Text);
                    cmd4.Parameters.AddWithValue("@Boy", textBox2.Text);
                    cmd4.Parameters.AddWithValue("@En", textBox6.Text);
                    cmd4.Parameters.AddWithValue("@Adet", textBox7.Text);
                    cmd4.Parameters.AddWithValue("@M2", textBox12.Text);
                    cmd4.Parameters.AddWithValue("@M2Fiyat", textBox15.Text);
                    cmd4.Parameters.AddWithValue("@Fiyat2", textBox17.Text);
                    cmd4.Parameters.AddWithValue("@AnaSiparişMi", "Evet");
                    cmd4.Parameters.AddWithValue("@İskonto", textBox16.Text);
                    cmd4.Parameters.AddWithValue("@EkleyenKullanici", kullaniciadi);

                    cmd4.ExecuteNonQuery();




                    İstatistik();
                    liste();
                    kayıtsayısı();
                    M2Toplat();
                    KapakAdetToplat();
                    TasarımUcretıToplat();
                    KapakToplam();
                    M2KapakSayısı();
                    if (comboBox2.Text == "Acil")
                    {
                        AcilHesaplama();
                    }
                    M2KapakFarkıHesaplama();
                    KargoHesaplama();
                    Toplam();
                    AraToplam();
                    DDSHesaplama();
                    GenelToplam();
                    StokGetir();
                    TutkalStok();
                    MDFStok();
                    temizle();
                }
                else
                {
                    MessageBox.Show("Ekleme yapabilmek için en, boy, adet bilgilerini yazın.", "Mesaj", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }
        private void M2Toplat()
        {
            try
            {
                toplamM2 = 0;
                if (dataGridView1.RowCount > 1)
                {
                    for (int i = 0; i < dataGridView1.Rows.Count; ++i)
                    {
                        if (dataGridView1.Rows[i].Cells["M2"].Value != null)
                        {
                            toplamM2 += Convert.ToDouble(dataGridView1.Rows[i].Cells["M2"].Value);
                        }
                    }
                    label27.Text = toplamM2.ToString("0.##");
                }
                else
                {

                }
            }
            catch (Exception)
            {

                throw;
            }


        }
        private void KapakAdetToplat()
        {
            try
            {
                toplamKapakAdet = 0;
                for (int i = 0; i < dataGridView1.Rows.Count; ++i)
                {
                    if (dataGridView1.Rows[i].Cells[7].Value != null)
                    {
                        toplamKapakAdet += Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value);
                    }
                }
                label28.Text = toplamKapakAdet.ToString("0.##");
            }
            catch (Exception)
            {

                throw;
            }


        }
        private void KapakToplam()
        {
            try
            {
                kapaktoplam = 0;
                for (int i = 0; i < dataGridView1.Rows.Count; ++i)
                {
                    if (dataGridView1.Rows[i].Cells[10].Value != null)
                    {
                        kapaktoplam += Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value);
                    }

                }
                label33.Text = kapaktoplam.ToString("0.##");
            }
            catch (Exception)
            {

            }


        }
        private void TasarımUcretıToplat()
        {
            try
            {
                toplamTasarımUcreti = 0;
                for (int i = 0; i < dataGridView1.Rows.Count; ++i)
                {
                    var cellValue = dataGridView1.Rows[i].Cells[3].Value;
                    if (cellValue != null && cellValue != DBNull.Value)
                    {
                        toplamTasarımUcreti += Convert.ToDouble(cellValue);
                    }
                }

                label34.Text = toplamTasarımUcreti.ToString("0.##");
            }
            catch (Exception ex)
            {
                // Hata ile ilgili bilgi almak isterseniz, ex nesnesini kullanabilirsiniz.
                MessageBox.Show("Bir hata oluştu: " + ex.Message);
            }



        }
        private void M2KapakSayısı()
        {
            try
            {
                m2kapaksayısı = 0;
                m2kapaksayısı += Convert.ToDouble(label28.Text) / Convert.ToDouble(label27.Text);
                label30.Text = m2kapaksayısı.ToString("0.##");
            }
            catch (Exception)
            {

                throw;
            }


        }
        double araToplam;
        private void AraToplam()
        {
            try
            {
                System.Globalization.CultureInfo culture = new System.Globalization.CultureInfo("tr-TR");
                araToplam = (Convert.ToDouble(label46.Text, culture) - Convert.ToDouble(label42.Text, culture) + Convert.ToDouble(label76.Text, culture));
                label74.Text = araToplam.ToString("N2", culture);

            }
            catch (Exception)
            {

            }
        }
        double toplam;
        private void Toplam()
        {
            try
            {
                System.Globalization.CultureInfo culture = new System.Globalization.CultureInfo("tr-TR");

                // Sayı biçimlendirmesini ayarlayın
                toplam = (Convert.ToDouble(label33.Text, culture) + Convert.ToDouble(label34.Text, culture) + Convert.ToDouble(label36.Text, culture) + Convert.ToDouble(label38.Text, culture));
                label46.Text = (toplam / (1.20)).ToString("N2", culture);
            }
            catch (Exception)
            {

            }
        }

        private void GenelToplam()
        {
            try
            {
                System.Globalization.CultureInfo culture = new System.Globalization.CultureInfo("tr-TR");

                geneltoplam = (araToplam + dds);
                label63.Text = geneltoplam.ToString("N2", culture);
            }
            catch (Exception)
            {

            }
        }
        double acilyüzde;
        double acil2;
        private void AcilHesaplama()
        {
            try
            {
                acil2 = 0;
                acilyüzde = 0;
                if (comboBox2.Text == "Acil")
                {
                    if (toplamM2 <= 0.5)
                    {
                        SqlCommand komut = new SqlCommand();
                        komut.CommandText = "SELECT *FROM Acil_Fiyatı where id=@p1";
                        komut.Parameters.AddWithValue("@p1", 1);
                        komut.Connection = bgl.baglanti();
                        komut.CommandType = CommandType.Text;

                        SqlDataReader dr;
                        dr = komut.ExecuteReader();
                        while (dr.Read())
                        {
                            acilyüzde = Convert.ToDouble(dr["Yüzde"]);
                        }
                        acil = kapaktoplam * (acilyüzde / 100);
                    }
                    else if (toplamM2 > 0.5 && toplamM2 <= 1)
                    {
                        SqlCommand komut = new SqlCommand();
                        komut.CommandText = "SELECT *FROM Acil_Fiyatı where id=@p1";
                        komut.Parameters.AddWithValue("@p1", 2);
                        komut.Connection = bgl.baglanti();
                        komut.CommandType = CommandType.Text;

                        SqlDataReader dr;
                        dr = komut.ExecuteReader();
                        while (dr.Read())
                        {
                            acilyüzde = Convert.ToDouble(dr["Yüzde"]);
                        }
                        acil = kapaktoplam * (acilyüzde / 100);
                    }
                    else if (toplamM2 > 1)
                    {
                        SqlCommand komut = new SqlCommand();
                        komut.CommandText = "SELECT *FROM Acil_Fiyatı where id=@p1";
                        komut.Parameters.AddWithValue("@p1", 3);
                        komut.Connection = bgl.baglanti();
                        komut.CommandType = CommandType.Text;

                        SqlDataReader dr;
                        dr = komut.ExecuteReader();
                        while (dr.Read())
                        {
                            acilyüzde = Convert.ToDouble(dr["Yüzde"]);
                        }
                        acil = kapaktoplam * (acilyüzde / 100);
                    }
                    label38.Text = acil.ToString("0.##");
                }
                else
                {
                    label38.Text = acil2.ToString("0.##");
                }
            }
            catch (Exception)
            {

                throw;
            }

        }
        double m2yüzde;
        double m2kapakfarkı1denkücükse;
        double m2kapakfarkı1denkücükseyüzde;
        private void M2KapakFarkıHesaplama()
        {
            try
            {
                m2yüzde = 0;
                m2kapakfarkı = 0;
                m2kapakfarkı1denkücükse = 0;
                m2kapakfarkı1denkücükseyüzde = 0;
                if (m2kapaksayısı >= 6 && m2kapaksayısı < 8)
                {
                    SqlCommand komut = new SqlCommand();
                    komut.CommandText = "SELECT *FROM M2_Kapak_Farkı where id=@p1";
                    komut.Parameters.AddWithValue("@p1", 1);
                    komut.Connection = bgl.baglanti();
                    komut.CommandType = CommandType.Text;

                    SqlDataReader dr;
                    dr = komut.ExecuteReader();
                    while (dr.Read())
                    {
                        m2yüzde = Convert.ToDouble(dr["Yüzde"]);
                    }
                    m2kapakfarkı = kapaktoplam * (m2yüzde / 100);
                }
                else if (m2kapaksayısı >= 8 && m2kapaksayısı < 10)
                {
                    SqlCommand komut = new SqlCommand();
                    komut.CommandText = "SELECT *FROM M2_Kapak_Farkı where id=@p1";
                    komut.Parameters.AddWithValue("@p1", 2);
                    komut.Connection = bgl.baglanti();
                    komut.CommandType = CommandType.Text;

                    SqlDataReader dr;
                    dr = komut.ExecuteReader();
                    while (dr.Read())
                    {
                        m2yüzde = Convert.ToDouble(dr["Yüzde"]);
                    }
                    m2kapakfarkı = kapaktoplam * (m2yüzde / 100);
                }
                else if (m2kapaksayısı >= 10)
                {
                    SqlCommand komut = new SqlCommand();
                    komut.CommandText = "SELECT *FROM M2_Kapak_Farkı where id=@p1";
                    komut.Parameters.AddWithValue("@p1", 3);
                    komut.Connection = bgl.baglanti();
                    komut.CommandType = CommandType.Text;

                    SqlDataReader dr;
                    dr = komut.ExecuteReader();
                    while (dr.Read())
                    {
                        m2yüzde = Convert.ToDouble(dr["Yüzde"]);
                    }
                    m2kapakfarkı = kapaktoplam * (m2yüzde / 100);
                }
                label36.Text = m2kapakfarkı.ToString("0.##");

                // M2 Kapak Farkı 1 Den Küçükse Kısmı
                if (toplamM2 < 1 && toplamM2 >= 0.5)
                {
                    SqlCommand komut = new SqlCommand();
                    komut.CommandText = "SELECT *FROM M2_Kapak_Farkı_2 where id=@p1";
                    komut.Parameters.AddWithValue("@p1", 1);
                    komut.Connection = bgl.baglanti();
                    komut.CommandType = CommandType.Text;

                    SqlDataReader dr;
                    dr = komut.ExecuteReader();
                    while (dr.Read())
                    {
                        m2kapakfarkı1denkücükseyüzde = Convert.ToDouble(dr["Yüzde"]);
                    }
                    m2kapakfarkı1denkücükse = kapaktoplam * (m2kapakfarkı1denkücükseyüzde / 100);
                }
                else if (toplamM2 < 0.5 && toplamM2 >= 0.25)
                {
                    SqlCommand komut = new SqlCommand();
                    komut.CommandText = "SELECT *FROM M2_Kapak_Farkı_2 where id=@p1";
                    komut.Parameters.AddWithValue("@p1", 2);
                    komut.Connection = bgl.baglanti();
                    komut.CommandType = CommandType.Text;

                    SqlDataReader dr;
                    dr = komut.ExecuteReader();
                    while (dr.Read())
                    {
                        m2kapakfarkı1denkücükseyüzde = Convert.ToDouble(dr["Yüzde"]);
                    }
                    m2kapakfarkı1denkücükse = kapaktoplam * (m2kapakfarkı1denkücükseyüzde / 100);
                }
                else if (toplamM2 < 0.25 && toplamM2 >= 0.1)
                {
                    SqlCommand komut = new SqlCommand();
                    komut.CommandText = "SELECT *FROM M2_Kapak_Farkı_2 where id=@p1";
                    komut.Parameters.AddWithValue("@p1", 3);
                    komut.Connection = bgl.baglanti();
                    komut.CommandType = CommandType.Text;

                    SqlDataReader dr;
                    dr = komut.ExecuteReader();
                    while (dr.Read())
                    {
                        m2kapakfarkı1denkücükseyüzde = Convert.ToDouble(dr["Yüzde"]);
                    }
                    m2kapakfarkı1denkücükse = kapaktoplam * (m2kapakfarkı1denkücükseyüzde / 100);
                }
                else if (toplamM2 < 0.1)
                {
                    SqlCommand komut = new SqlCommand();
                    komut.CommandText = "SELECT *FROM M2_Kapak_Farkı_2 where id=@p1";
                    komut.Parameters.AddWithValue("@p1", 4);
                    komut.Connection = bgl.baglanti();
                    komut.CommandType = CommandType.Text;

                    SqlDataReader dr;
                    dr = komut.ExecuteReader();
                    while (dr.Read())
                    {
                        m2kapakfarkı1denkücükseyüzde = Convert.ToDouble(dr["Yüzde"]);
                    }
                    m2kapakfarkı1denkücükse = kapaktoplam * (m2kapakfarkı1denkücükseyüzde / 100);
                }
                label36.Text = (m2kapakfarkı1denkücükse + m2kapakfarkı).ToString("0.##");
            }
            catch (Exception)
            {

                throw;
            }


        }
        double kargofiyati;
        double kargofiyatikucukse;
        double kargofiyatikucukkucukse;

        private void KargoHesaplama()
        {
            try
            {
                kargofiyati = 0;
                if (comboBox3.Text == "Kargo")
                {
                    SqlCommand komut = new SqlCommand();
                    komut.CommandText = "SELECT *FROM Kargo_Fiyatı where id=@p1";
                    komut.Parameters.AddWithValue("@p1", 1);
                    komut.Connection = bgl.baglanti();
                    komut.CommandType = CommandType.Text;

                    SqlDataReader dr;
                    dr = komut.ExecuteReader();
                    while (dr.Read())
                    {
                        kargofiyati = Convert.ToDouble(dr["KargoFiyatı"]);
                        kargofiyatikucukse = Convert.ToDouble(dr["KargoFiyatiKucukse"]);
                        kargofiyatikucukkucukse = Convert.ToDouble(dr["KargoFiyatiKucukKucukse"]);

                    }

                    if (toplamM2 >= 0.5 && toplamM2 <= 1)
                    {
                        kargo = kargofiyatikucukse;
                    }
                    else if (toplamM2 > 0 && toplamM2 < 0.5)
                    {
                        kargo = kargofiyatikucukkucukse;
                    }
                    else
                    {
                        //1METREKARE ÜSTÜ
                        kargo = toplamM2 * kargofiyati;
                    }

                    label40.Text = kargo.ToString("0.##");
                    label76.Text = kargo.ToString("0.##");
                }
                else
                {
                    label40.Text = kargofiyati.ToString("0.##");
                    label76.Text = kargofiyati.ToString("0.##");
                }
            }
            catch (Exception)
            {

                throw;
            }



        }
        private void DDSHesaplama()
        {
            try
            {
                dds = 0;
                dds = (araToplam * 20) / 100;
                //dds = geneltoplam - (geneltoplam / (1.2));
                label44.Text = dds.ToString("0.##");
            }
            catch (Exception)
            {

                throw;
            }


        }
        private void SiparişBölmeSayısı()
        {
            try
            {
                SqlCommand cmd = new SqlCommand("select count(DISTINCT Renk) Renk from Siparişler Where SiparisNo=@p1", bgl.baglanti());
                cmd.Parameters.AddWithValue("@p1", textBox1.Text);
                bölmekayitsayisi = Convert.ToInt32(cmd.ExecuteScalar());
            }
            catch (Exception)
            {

                throw;
            }

        }
        private void SiparişBölme()
        {
            try
            {
                SiparişBölmeSayısı();
                if (bölmekayitsayisi > 1)
                {
                    SqlCommand komut = new SqlCommand();
                    komut.CommandText = "select DISTINCT Renk from Siparişler where SiparisNo=@p1";
                    komut.Parameters.AddWithValue("@p1", textBox1.Text);
                    komut.Connection = bgl.baglanti();
                    komut.CommandType = CommandType.Text;

                    SqlDataReader dr;
                    dr = komut.ExecuteReader();
                    while (dr.Read())
                    {
                        comboBox8.Items.Add(dr["Renk"].ToString());
                        bölmesayısı++;
                        message += "\n" + textBox1.Text + "-" + bölmesayısı + " " + dr["Renk"].ToString() + "\n" + " ";

                    }
                    DialogResult dialogResult = MessageBox.Show("Sipariş birden fazla siparişe bölünecek\n " + "" + message, "Bilgi", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        siparisbölme = "böl";
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        siparisbölme = "bölme";
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }



        }
        double stok;
        string birim;
        double tutkalstok;
        double mdfstok;
        private void StokGetir()
        {
            try
            {
                stok = 0;
                if (comboBox6.Text != "")
                {

                    SqlCommand komut = new SqlCommand();
                    komut.CommandText = "SELECT * FROM Stoklar where Malzeme=@Malzeme";
                    komut.Parameters.AddWithValue("@Malzeme", comboBox6.Text);
                    komut.Connection = bgl.baglanti();
                    komut.CommandType = CommandType.Text;

                    SqlDataReader dr;
                    dr = komut.ExecuteReader();
                    while (dr.Read())
                    {
                        stok = Convert.ToDouble(dr["Kalan"].ToString());
                        birim = dr["Birim"].ToString();
                    }
                    label55.Text = stok.ToString("0.##") + birim;

                }

                renkstok = 0;
                SqlCommand komut2 = new SqlCommand();
                komut2.CommandText = "SELECT * FROM Stoklar where Malzeme=@Malzeme";
                komut2.Parameters.AddWithValue("@Malzeme", comboBox6.Text);
                komut2.Connection = bgl.baglanti();
                komut2.CommandType = CommandType.Text;

                SqlDataReader dr2;
                dr2 = komut2.ExecuteReader();
                while (dr2.Read())
                {
                    renkstok = Convert.ToDouble(dr2["Kalan"].ToString());
                }

                label71.Text = comboBox6.Text + " = " + renkstok.ToString("0.##") + " m²";
            }
            catch (Exception)
            {

                throw;
            }
        }
        private void TutkalStok()
        {
            try
            {
                tutkalstok = 0;

                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT * FROM Stoklar where Malzeme=@Malzeme";
                komut.Parameters.AddWithValue("@Malzeme", "TUTKAL");
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    tutkalstok = Convert.ToDouble(dr["Kalan"].ToString());
                }

                label51.Text = tutkalstok.ToString("0.##");
            }
            catch (Exception)
            {

                throw;
            }
        }
        private void MDFStok()
        {
            try
            {
                mdfstok = 0;
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT * FROM Stoklar where Malzeme=@Malzeme";
                komut.Parameters.AddWithValue("@Malzeme", "18 MM TEKYÜZ MDF");
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    mdfstok = Convert.ToDouble(dr["Kalan"].ToString());
                }

                label49.Text = mdfstok.ToString("0.##");
            }
            catch (Exception)
            {

                throw;
            }

        }
        public void groupBox1_Paint(object sender, PaintEventArgs e)
        {
            System.Windows.Forms.GroupBox box = sender as System.Windows.Forms.GroupBox;
            DrawGroupBox(box, e.Graphics, Color.Black, Color.Black);
        }
        public void DrawGroupBox(System.Windows.Forms.GroupBox box, Graphics g, Color textColor, Color borderColor)
        {
            if (box != null)
            {
                Brush textBrush = new SolidBrush(textColor);
                Brush borderBrush = new SolidBrush(borderColor);
                Pen borderPen = new Pen(borderBrush);
                SizeF strSize = g.MeasureString(box.Text, box.Font);
                System.Drawing.Rectangle rect = new System.Drawing.Rectangle(box.ClientRectangle.X,
                                               box.ClientRectangle.Y + (int)(strSize.Height / 2),
                                               box.ClientRectangle.Width - 1,
                                               box.ClientRectangle.Height - (int)(strSize.Height / 2) - 1);
                // Clear text and border
                g.Clear(this.BackColor);
                // Draw text
                g.DrawString(box.Text, box.Font, textBrush, box.Padding.Left, 0);
                // Drawing Border
                //Left
                g.DrawLine(borderPen, rect.Location, new System.Drawing.Point(rect.X, rect.Y + rect.Height));
                //Right
                g.DrawLine(borderPen, new System.Drawing.Point(rect.X + rect.Width, rect.Y), new System.Drawing.Point(rect.X + rect.Width, rect.Y + rect.Height));
                //Bottom
                g.DrawLine(borderPen, new System.Drawing.Point(rect.X, rect.Y + rect.Height), new System.Drawing.Point(rect.X + rect.Width, rect.Y + rect.Height));
                //Top1
                g.DrawLine(borderPen, new System.Drawing.Point(rect.X, rect.Y), new System.Drawing.Point(rect.X + box.Padding.Left, rect.Y));
                //Top2
                g.DrawLine(borderPen, new System.Drawing.Point(rect.X + box.Padding.Left + (int)(strSize.Width), rect.Y), new System.Drawing.Point(rect.X + rect.Width, rect.Y));
            }
        }
        private bool formLoaded = false;
        private async Task EskiModeksAyrinti()
        {
            string sorgu = "SELECT SIPARISNO,Model,Renk,OZELLIK,Boy,En,Adet,M2,M2FIYATI,TUTARI FROM SIPAYRINTI WHERE SIPARISNO LIKE '" + SiparisNoSuEski + "%'";
            DataSet ds = new DataSet();

            await Task.Run(() =>
            {
                using (SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti_eski()))
                {
                    adap.Fill(ds, "Siparisler");
                }
            });


            dataGridView1.DataSource = ds.Tables["Siparisler"];
        }
        private async void Form10_Load(object sender, EventArgs e)
        {
            checkBox3.Checked = true;
            if (vbkontrol && yetki != "Yönetici")
            {
                groupBox4.Visible = false;
                groupBox3.Visible = false;
                button1.Visible = false;
                button3.Visible = false;
                button5.Visible = false;
                button6.Visible = false;
                button4.Visible = false;
                button13.Visible = false;
            } else
            {
                button14.Visible = false;
            }
            Show();



            Timer timer = new Timer();
            timer.Interval = 100;
            timer.Tick += (s, args) =>
            {
                timer.Stop();
                Asm();
            };


            timer.Start();

        }
        private Timer timer;

        public async void Asm()
        {
            if (hangiformdanModeksEski == "true")
            {
                checkBox1.Visible = false;
                groupBox3.Visible = false;
                button1.Enabled = false;
                button3.Enabled = false;
                button4.Enabled = false;
                await EskiModeksAyrinti();
                comboBox1.Enabled = false;
                comboBox2.Enabled = false;
                comboBox3.Enabled = false;
                comboBox5.Enabled = false;
                comboBox6.Enabled = false;
                comboBox7.Enabled = false;
                textBox3.Enabled = false;
                textBox4.Enabled = false;
                textBox5.Enabled = false;
                textBox8.Enabled = false;
                textBox10.Enabled = false;
                textBox11.Enabled = false;
                textBox13.Enabled = false;
                textBox14.Enabled = false;
                textBox15.Enabled = false;
                textBox19.Enabled = false;
                button8.Enabled = false;
                button9.Enabled = false;
                button10.Enabled = false;
                dateTimePicker1.Enabled = false;
                textBox1.Text = SiparisNoSuEski;
                comboBox1.Text = MusteriEski;
                dateTimePicker1.Value = Convert.ToDateTime(SiparisTarihi);
                comboBox4.DropDownStyle = ComboBoxStyle.DropDown;
                comboBox5.DropDownStyle = ComboBoxStyle.DropDown;
                comboBox6.DropDownStyle = ComboBoxStyle.DropDown;
                comboBox4.Text = dataGridView1.Rows[0].Cells["Model"].Value.ToString();
                comboBox6.Text = dataGridView1.Rows[0].Cells["Renk"].Value.ToString();
                textBox15.Text = dataGridView1.Rows[0].Cells["M2FIYATI"].Value.ToString();
                comboBox7.Text = dataGridView1.Rows[0].Cells["OZELLIK"].Value.ToString();
                textBox2.Text = dataGridView1.Rows[0].Cells["Boy"].Value.ToString();
                textBox6.Text = dataGridView1.Rows[0].Cells["En"].Value.ToString();
                textBox7.Text = dataGridView1.Rows[0].Cells["Adet"].Value.ToString();
                textBox12.Text = dataGridView1.Rows[0].Cells["M2"].Value.ToString();
                textBox17.Text = dataGridView1.Rows[0].Cells["TUTARI"].Value.ToString();
            }
            else
            {
                label59.Text = kullaniciadi;
                formLoaded = true;
                müsteri_cek();
                satıs_ıd_getir();
                checkBox2.Checked = false;
                groupBox5.Visible = false;
                textBox1.Text = siparişno.ToString();
                label56.Text = siparişno.ToString();
                modelbilgilerigetir();
                renkbilgilerigetir();
                özellikbilgilerigetir();
                kayıtsayısı();
                if (textBox1.Text != "")
                {
                    textBox1.Text = siparişno;

                    SqlCommand komut = new SqlCommand();
                    komut.CommandText = "SELECT *FROM Siparişler where SiparisNo=@p1 AND AnaSiparişMi=@p2";
                    komut.Parameters.AddWithValue("@p1", textBox1.Text);
                    komut.Parameters.AddWithValue("@p2", "Evet");

                    komut.Connection = bgl.baglanti();
                    komut.CommandType = CommandType.Text;

                    SqlDataReader dr;
                    dr = komut.ExecuteReader();
                    while (dr.Read())
                    {
                        comboBox2.SelectedItem = dr["SiparişTipi"];
                        comboBox3.SelectedItem = dr["SevkTürü"];
                        textBox11.Text = dr["Nott"].ToString();
                        textBox19.Text = dr["OnayTarihi"].ToString();
                        dateTimePicker1.Text = dr["SiparişTarihi"].ToString();
                        textBox10.Text = dr["TeslimTarihi"].ToString();
                        //textBox11.Enabled = false;
                        if (dr["Onay"].ToString() == "Onaylandı")
                        {
                            comboBox2.SelectedItem = dr["SiparişTipi"];
                            comboBox3.SelectedItem = dr["SevkTürü"];
                            comboBox1.Enabled = false;
                            comboBox2.Enabled = false;
                            comboBox3.Enabled = false;
                            checkBox1.Visible = false;
                            button9.Visible = false;
                            button10.Visible = false;
                            button8.Visible = false;
                            button1.Enabled = false;
                            button2.Enabled = false;
                            button3.Enabled = false;
                            button11.Enabled = false;
                            if (Convert.ToDouble(dr["İskonto"]) > 0)
                            {
                                label58.Text = "İskonto Yapıldı";
                            }
                            else if (Convert.ToDouble(dr["İskonto"]) == 0)
                            {
                                label58.Text = "İskonto Yapılmadı";
                            }
                            textBox16.Enabled = false;
                            if (button3.Enabled == false)
                            {
                                button3.BackColor = Color.Green;
                            }
                            button3.Text = "ONAYLANDI";
                            textBox9.Text = dr["SiparişTarihi"].ToString();
                            textBox10.Text = dr["TeslimTarihi"].ToString();

                        }
                        label67.Text = kullaniciadi;
                        resimgetir();
                        checkBox3.Checked = true;
                        birlesikliste();

                    }
                }
                else if (textBox1.Text == "")
                {
                    if (kayitSayisi < 100)
                    {
                        if (kayitSayisi == 99)
                        {
                            textBox1.Text = (kayitSayisi + 101).ToString();
                        }
                        else
                        {
                            textBox1.Text = (kayitSayisi + 100).ToString();
                        }
                    }
                    else
                    {
                        textBox1.Text = (kayitSayisi + 1).ToString();
                    }
                }

                textBox2.Text = boy.ToString();
                textBox6.Text = en.ToString();
                textBox7.Text = adet.ToString();
                textBox12.Text = m2.ToString("0.##");

                timer1.Start();
                liste();
                if (button2.Enabled == true)
                {
                    //textBox9.Text = DateTime.Now.ToShortDateString();
                    //textBox10.Text = DateTime.Now.AddDays(20).ToString("dd.MM.yyyy HH:dd:ss");
                }
                M2Toplat();
                KapakAdetToplat();
                TasarımUcretıToplat();
                KapakToplam();
                M2KapakSayısı();
                if (comboBox2.Text == "Acil")
                {
                    AcilHesaplama();
                }
                M2KapakFarkıHesaplama();
                KargoHesaplama();
                Toplam();
                AraToplam();
                DDSHesaplama();
                GenelToplam();
                if (textBox2.Text != "" && textBox6.Text != "" && textBox7.Text != "")
                {
                    formül();
                }
                TutkalStok();
                MDFStok();
                iskontogetir();


                Toplam();
                AraToplam();
                GenelToplam();
                label67.Text = kullaniciadi;

                SiparişBölmeSayısı();
                if (bölmekayitsayisi > 1)
                    checkBox2.Enabled = false;
                else
                    checkBox2.Enabled = true;

                resimgetir();

            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            label2.Text = DateTime.Now.ToLongDateString();
            label12.Text = DateTime.Now.ToLongTimeString();
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar) && e.KeyChar != ',';
            if (e.KeyChar == (char)Keys.Enter)
            {
                fiyatDeğişiklikEklemeKontrol();
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (comboBox2.Text == "Acil" || comboBox2.Text == "Üretim Sorunu")
                {
                    textBox10.Text = dateTimePicker1.Value.AddDays(5).ToString("dd.MM.yyyy HH:dd:ss");
                    //liste();
                    M2Toplat();
                    KapakAdetToplat();
                    TasarımUcretıToplat();
                    KapakToplam();
                    M2KapakSayısı();
                    AcilHesaplama();
                    M2KapakFarkıHesaplama();
                    KargoHesaplama();
                    Toplam();
                    AraToplam();
                    DDSHesaplama();
                    GenelToplam();

                    string sorgu = "UPDATE Siparişler SET AcilFarkı=@AcilFarkı WHERE SiparisNo=@SiparisNo";
                    SqlCommand komut;
                    komut = new SqlCommand(sorgu, bgl.baglanti());
                    komut.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
                    komut.Parameters.AddWithValue("@AcilFarkı", label38.Text);
                    komut.ExecuteNonQuery();

                    string sorgu2 = "UPDATE Siparişler SET SiparişTipi=@SiparişTipi WHERE SiparisNo=@SiparisNo";
                    SqlCommand komut2;
                    komut2 = new SqlCommand(sorgu2, bgl.baglanti());
                    komut2.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
                    komut2.Parameters.AddWithValue("@SiparişTipi", comboBox2.Text);
                    komut2.ExecuteNonQuery();

                }
                else if (comboBox2.Text == "Normal")
                {
                    textBox10.Text = dateTimePicker1.Value.AddDays(20).ToString("dd.MM.yyyy HH:dd:ss");
                    //liste();
                    M2Toplat();
                    KapakAdetToplat();
                    TasarımUcretıToplat();
                    KapakToplam();
                    M2KapakSayısı();
                    AcilHesaplama();
                    M2KapakFarkıHesaplama();
                    KargoHesaplama();
                    Toplam();
                    AraToplam();
                    DDSHesaplama();
                    GenelToplam();
                   

                    string sorgu = "UPDATE Siparişler SET AcilFarkı=@AcilFarkı WHERE SiparisNo=@SiparisNo";
                    SqlCommand komut;
                    komut = new SqlCommand(sorgu, bgl.baglanti());
                    komut.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
                    komut.Parameters.AddWithValue("@AcilFarkı", label38.Text);
                    komut.ExecuteNonQuery();

                    string sorgu2 = "UPDATE Siparişler SET SiparişTipi=@SiparişTipi WHERE SiparisNo=@SiparisNo";
                    SqlCommand komut2;
                    komut2 = new SqlCommand(sorgu2, bgl.baglanti());
                    komut2.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
                    komut2.Parameters.AddWithValue("@SiparişTipi", comboBox2.Text);
                    komut2.ExecuteNonQuery();
                }
            }
            catch (Exception)
            {

                throw;
            }

        }

        private void textBox2_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                if (textBox2.Text != "")
                {
                    boy = float.Parse(textBox2.Text);
                    if (boy > 250)
                    {
                        MessageBox.Show("Boy 250'den büyük olamaz.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        boy = 0;
                        textBox2.Text = boy.ToString();
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Doğru bir biçim giriniz.");
            }

            // Yalnızca bir ondalık basamağa izin ver
            string[] parts = textBox2.Text.Split(',');
            if (parts.Length > 1 && parts[1].Length > 1)
            {
                textBox2.Text = parts[0] + ',' + parts[1][0]; // Sadece bir ondalık basamağı al
                textBox2.SelectionStart = textBox2.Text.Length; // Cursor'ı metnin sonuna taşı
            }

            if (textBox2.Text != "" && textBox6.Text != "" && textBox7.Text != "")
            {
                formül();
            }

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (textBox6.Text != "")
                {
                    en = float.Parse(textBox6.Text);
                    if (en > 120)
                    {
                        MessageBox.Show("En 120'den büyük olamaz.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        en = 0;
                        textBox6.Text = en.ToString();
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Doğru bir biçim giriniz.");
            }

            // Yalnızca bir ondalık basamağa izin ver
            string[] parts = textBox6.Text.Split(',');
            if (parts.Length > 1 && parts[1].Length > 1)
            {
                textBox6.Text = parts[0] + ',' + parts[1][0]; // Sadece bir ondalık basamağı al
                textBox6.SelectionStart = textBox6.Text.Length; // Cursor'ı metnin sonuna taşı
            }
            if (textBox2.Text != "" && textBox6.Text != "" && textBox7.Text != "")
            {
                formül();
            }

        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            if (textBox2.Text == "0")
            {
                textBox2.Text = "";
            }

        }

        private void textBox6_Click(object sender, EventArgs e)
        {
            if (textBox6.Text == "0")
            {
                textBox6.Text = "";
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox13.Text = textBox1.Text;
        }

        private void textBox7_Click(object sender, EventArgs e)
        {
            if (textBox7.Text == "0")
            {
                textBox7.Text = "";
            }
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (textBox7.Text != "")
                {
                    adet = float.Parse(textBox7.Text);
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Doğru bir biçim giriniz.");
            }


            if (textBox2.Text != "" && textBox6.Text != "" && textBox7.Text != "")
            {
                formül();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            müsteri_bilgileri_getir();
        }

        private void ekle()
        {
            try
            {
                kaydet();
                boy = 0;
                en = 0;
                adet = 0;
                m2 = 0;
                bölmesayısı = 0;
                message = "";
                textBox2.Text = boy.ToString();
                textBox6.Text = en.ToString();
                textBox7.Text = adet.ToString();
                textBox12.Text = m2.ToString("0.##");

                fiyatbilgisigetir();
                hgfiyatgetir();
                RSMfiyatgetir();
                SOFTfiyatgetir();
                ULTRASOFTfiyatgetir();

                string sorgu = "UPDATE Siparişler SET ToplamM2=@ToplamM2,ToplamAdet=@ToplamAdet,ToplamFiyat=@ToplamFiyat,ToplamTasarımÜcreti=@ToplamTasarımÜcreti,İskonto=@İskonto,Kargo=@Kargo,AcilFarkı=@AcilFarkı,M2KapakFarkı=@M2KapakFarkı,M2KapakAdet=@M2KapakAdet,DDS=@DDS, Aşama = @Aşama WHERE SiparisNo=@SiparisNo AND AnaSiparişMi=@AnaSiparişMi";
                SqlCommand komut;
                komut = new SqlCommand(sorgu, bgl.baglanti());
                komut.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
                komut.Parameters.AddWithValue("@AnaSiparişMi", "Evet");
                komut.Parameters.AddWithValue("@Aşama", "Onay Bekliyor");
                komut.Parameters.AddWithValue("@ToplamM2", label27.Text);
                komut.Parameters.AddWithValue("@ToplamAdet", label28.Text);
                komut.Parameters.AddWithValue("@ToplamFiyat", label63.Text);
                komut.Parameters.AddWithValue("@ToplamTasarımÜcreti", label34.Text);
                komut.Parameters.AddWithValue("@İskonto", textBox16.Text);
                komut.Parameters.AddWithValue("@Kargo", label40.Text);
                komut.Parameters.AddWithValue("@AcilFarkı", label38.Text);
                komut.Parameters.AddWithValue("@M2KapakFarkı", label36.Text);
                komut.Parameters.AddWithValue("@M2KapakAdet", label30.Text);
                komut.Parameters.AddWithValue("@DDS", dds.ToString("0.##"));
                komut.ExecuteNonQuery();

                SiparişBölmeSayısı();
                if (bölmekayitsayisi > 1)
                    checkBox2.Enabled = false;
                else
                    checkBox2.Enabled = true;


                birlesikliste();
            }
            catch (Exception)
            {

                throw;
            }


        }
        private void fiyatDeğişiklikEklemeKontrol()
        {
            try
            {
                if (textBox15.Text != fiyat1.ToString() || textBox15.Text != fiyat1.ToString())
                {
                    if (textBox11.Text.Length > 5)
                    {
                        DialogResult d1 = new DialogResult();
                        d1 = MessageBox.Show(comboBox6.Text + " " + "renginin fiyatını değiştirdiniz, Kaydetmek istiyor musunuz?", "Bilgi", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                        if (d1 == DialogResult.Yes)
                        {
                            ekle();
                        }
                    }
                    else if (textBox11.Text.Length < 5)
                    {
                        MessageBox.Show(comboBox6.Text + " renginin fiyatını değiştirdiniz, lütfen not alanına bir açıklama yapın.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }
                    else if (textBox15.Text == "0")
                    {
                        if (textBox11.Text.Length < 5)
                        {
                            MessageBox.Show("Fiyatı 0 girdiniz. Lütfen bir açıklama giriniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else
                        {
                            DialogResult d2 = new DialogResult();
                            d2 = MessageBox.Show(comboBox6.Text + " " + "renginin fiyatını 0 girdiniz, Kaydetmek istiyor musunuz?", "Bilgi", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                        }
                    }
                }
                else
                {
                    ekle();
                }
            }
            catch (Exception)
            {

                throw;
            }

        }
        private void button8_Click(object sender, EventArgs e)
        {
            fiyatDeğişiklikEklemeKontrol();
        }

        private void veridoldur()
        {
            try
            {
                if (checkBox1.Checked == true)
                {
                    comboBox2.Enabled = true;
                    comboBox3.Enabled = true;
                    id = dataGridView1.CurrentRow.Cells[0].Value.ToString();
                    SqlCommand komut = new SqlCommand();
                    komut.CommandText = "SELECT *FROM Siparişler where id=@id";
                    komut.Parameters.AddWithValue("@id", id);
                    komut.Connection = bgl.baglanti();
                    komut.CommandType = CommandType.Text;

                    SqlDataReader dr;
                    dr = komut.ExecuteReader();
                    while (dr.Read())
                    {
                        comboBox1.Text = dr["Müşteri"].ToString();
                        comboBox2.Text = dr["SiparişTipi"].ToString();
                        comboBox3.Text = dr["SevkTürü"].ToString();
                        textBox8.Text = dr["BID"].ToString();
                        textBox9.Text = dr["SiparişTarihi"].ToString();
                        textBox10.Text = dr["TeslimTarihi"].ToString();
                        textBox11.Text = dr["Nott"].ToString();
                        comboBox4.Text = dr["Model"].ToString();
                        comboBox6.Text = dr["Renk"].ToString();
                        textBox15.Text = dr["M2Fiyat"].ToString();
                        textBox14.Text = dr["TasarımÜcreti"].ToString();
                        comboBox7.Text = dr["Özellik"].ToString();
                        if (dr["Özellik"].ToString() == "")
                        {
                            comboBox7.Text = " ";
                        }
                        comboBox5.Text = dr["BaskıYönü"].ToString();
                        textBox2.Text = dr["Boy"].ToString();
                        textBox6.Text = dr["En"].ToString();
                        textBox7.Text = dr["Adet"].ToString();
                        textBox12.Text = dr["M2"].ToString();
                        textBox17.Text = dr["Fiyat2"].ToString();
                        label73.Text = dr["TasarımÜcreti"].ToString();
                    }
                }
                else
                {
                    //comboBox2.Enabled = false;
                    //comboBox3.Enabled = false;
                    id = dataGridView1.CurrentRow.Cells[0].Value.ToString();
                    SqlCommand komut = new SqlCommand();
                    komut.CommandText = "SELECT *FROM Siparişler where id=@id";
                    komut.Parameters.AddWithValue("@id", id);
                    komut.Connection = bgl.baglanti();
                    komut.CommandType = CommandType.Text;

                    SqlDataReader dr;
                    dr = komut.ExecuteReader();
                    while (dr.Read())
                    {
                        comboBox1.Text = dr["Müşteri"].ToString();
                        comboBox4.Text = dr["Model"].ToString();
                        comboBox6.Text = dr["Renk"].ToString();
                        textBox8.Text = dr["BID"].ToString();
                    }
                    bgl.baglanti().Close();
                }
            }
            catch (Exception ex)
            {
                return;
            }
        }
        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (hangiformdanModeksEski != "true")
                veridoldur();
        }

        private void button10_Click(object sender, EventArgs e)

        {
            try
            {
                if (id != "")
                {
                    string sorgu = "DELETE FROM Siparişler WHERE id=@id";
                    SqlCommand komut;
                    komut = new SqlCommand(sorgu, bgl.baglanti());
                    komut.Parameters.AddWithValue("@id", id);
                    komut.ExecuteNonQuery();
                    MessageBox.Show("Kayıt başarıyla silindi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    liste();
                    M2Toplat();
                    KapakAdetToplat();
                    TasarımUcretıToplat();
                    KapakToplam();
                    M2KapakSayısı();
                    if (comboBox2.Text == "Acil")
                    {
                        AcilHesaplama();
                    }
                    M2KapakFarkıHesaplama();
                    KargoHesaplama();
                    Toplam();
                    AraToplam();
                    DDSHesaplama();
                    GenelToplam();

                    string sorgu2 = "UPDATE Siparişler SET ToplamM2=@ToplamM2,ToplamAdet=@ToplamAdet,ToplamFiyat=@ToplamFiyat,ToplamTasarımÜcreti=@ToplamTasarımÜcreti,İskonto=@İskonto,Kargo=@Kargo,AcilFarkı=@AcilFarkı,M2KapakFarkı=@M2KapakFarkı WHERE SiparisNo=@SiparisNo AND AnaSiparişMi=@AnaSiparişMi";
                    SqlCommand komut2;
                    komut2 = new SqlCommand(sorgu2, bgl.baglanti());
                    komut2.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
                    komut2.Parameters.AddWithValue("@AnaSiparişMi", "Evet");
                    komut2.Parameters.AddWithValue("@ToplamM2", label27.Text);
                    komut2.Parameters.AddWithValue("@ToplamAdet", label28.Text);
                    komut2.Parameters.AddWithValue("@ToplamFiyat", label63.Text);
                    komut2.Parameters.AddWithValue("@ToplamTasarımÜcreti", label34.Text);
                    komut2.Parameters.AddWithValue("@İskonto", textBox16.Text);
                    komut2.Parameters.AddWithValue("@Kargo", label40.Text);
                    komut2.Parameters.AddWithValue("@AcilFarkı", label38.Text);
                    komut2.Parameters.AddWithValue("@M2KapakFarkı", label36.Text);
                    komut2.ExecuteNonQuery();
                }
                else
                {
                    MessageBox.Show("Bir kayıt seçiniz.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception)
            {
                return;
            }

        }
        private void Güncelle()
        {
            try
            {
                int rowIndex = dataGridView1.CurrentCell.RowIndex;
                int id = Convert.ToInt32(dataGridView1.Rows[rowIndex].Cells[0].Value.ToString());

                MessageBox.Show(id.ToString());
                string sorgu = "UPDATE Siparişler SET Müşteri=@Müşteri,SiparişTipi=@SiparişTipi,SevkTürü=@SevkTürü,BID=@BID,Nott=@Nott,Model=@Model,Renk=@Renk,Fiyat=@Fiyat,TasarımÜcreti=@TasarımÜcreti,Özellik=@Özellik,BaskıYönü=@BaskıYönü,Boy=@Boy,En=@En,Adet=@Adet,M2=@M2,M2Fiyat=@M2Fiyat,Fiyat2=@Fiyat2 WHERE id="+ id + "";
                SqlCommand komut;
                komut = new SqlCommand(sorgu, bgl.baglanti());
                komut.Parameters.AddWithValue("@Müşteri", comboBox1.Text);
                komut.Parameters.AddWithValue("@SiparişTipi", comboBox2.Text);
                komut.Parameters.AddWithValue("@SevkTürü", comboBox3.Text);
                komut.Parameters.AddWithValue("@BID", textBox8.Text);
                komut.Parameters.AddWithValue("@Nott", textBox11.Text);
                komut.Parameters.AddWithValue("@Model", comboBox4.Text);
                komut.Parameters.AddWithValue("@Renk", comboBox6.Text);
                komut.Parameters.AddWithValue("@Fiyat", textBox15.Text);
                komut.Parameters.AddWithValue("@TasarımÜcreti", textBox14.Text);
                komut.Parameters.AddWithValue("@Özellik", comboBox7.Text);
                komut.Parameters.AddWithValue("@BaskıYönü", comboBox5.Text);
                komut.Parameters.AddWithValue("@Boy", textBox2.Text);
                komut.Parameters.AddWithValue("@En", textBox6.Text);
                komut.Parameters.AddWithValue("@Adet", textBox7.Text);
                komut.Parameters.AddWithValue("@M2", textBox12.Text);
                komut.Parameters.AddWithValue("@M2Fiyat", textBox15.Text);
                komut.Parameters.AddWithValue("@Fiyat2", Convert.ToString((Convert.ToDouble(textBox15.Text) * Convert.ToDouble(textBox12.Text))));
                komut.ExecuteNonQuery();
                MessageBox.Show("Kayıt başarıyla güncellendi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                liste();
                M2Toplat();
                KapakAdetToplat();
                TasarımUcretıToplat();
                KapakToplam();
                M2KapakSayısı();
                if (comboBox2.Text == "Acil")
                {
                    AcilHesaplama();
                }
                M2KapakFarkıHesaplama();
                KargoHesaplama();
                Toplam();
                AraToplam();
                DDSHesaplama();
                GenelToplam();

                string sorgu2 = "UPDATE Siparişler SET ToplamM2=@ToplamM2,ToplamAdet=@ToplamAdet,ToplamFiyat=@ToplamFiyat,ToplamTasarımÜcreti=@ToplamTasarımÜcreti,İskonto=@İskonto,Kargo=@Kargo,AcilFarkı=@AcilFarkı,M2KapakFarkı=@M2KapakFarkı,DDS=@DDS WHERE SiparisNo=@SiparisNo AND AnaSiparişMi=@AnaSiparişMi";
                SqlCommand komut2;
                komut2 = new SqlCommand(sorgu2, bgl.baglanti());
                komut2.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
                komut2.Parameters.AddWithValue("@AnaSiparişMi", "Evet");
                komut2.Parameters.AddWithValue("@ToplamM2", label27.Text);
                komut2.Parameters.AddWithValue("@ToplamAdet", label28.Text);
                komut2.Parameters.AddWithValue("@ToplamFiyat", label63.Text);
                komut2.Parameters.AddWithValue("@ToplamTasarımÜcreti", label34.Text);
                komut2.Parameters.AddWithValue("@İskonto", textBox16.Text);
                komut2.Parameters.AddWithValue("@Kargo", label40.Text);
                komut2.Parameters.AddWithValue("@AcilFarkı", label38.Text);
                komut2.Parameters.AddWithValue("@M2KapakFarkı", label36.Text);
                komut2.Parameters.AddWithValue("@DDS", dds.ToString());
                komut2.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                if (checkBox1.Checked == true && (textBox15.Text != fiyat1.ToString() || textBox15.Text != fiyat1.ToString()))
                {
                    if (textBox11.Text.Length > 5)
                    {
                        DialogResult d1 = new DialogResult();
                        d1 = MessageBox.Show(comboBox6.Text + " " + "renginin fiyatını değiştirdiniz, Güncellemek istiyor musunuz?", "Bilgi", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                        if (d1 == DialogResult.Yes)
                        {
                            Güncelle();
                            checkBox3.Checked = true;
                        }
                    }
                    else if (textBox11.Text.Length < 5)
                    {
                        MessageBox.Show(comboBox6.Text + " renginin fiyatını değiştirdiniz, lütfen not alanına bir açıklama yapın.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }
                    else if (textBox15.Text == "0")
                    {
                        if (textBox11.Text.Length < 5)
                        {
                            MessageBox.Show("Fiyatı 0 girdiniz. Lütfen bir açıklama giriniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else
                        {
                            DialogResult d2 = new DialogResult();
                            d2 = MessageBox.Show(comboBox6.Text + " " + "renginin fiyatını 0 girdiniz, Güncellemek istiyor musunuz?", "Bilgi", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                            if (d2 == DialogResult.Yes)
                            {
                                Güncelle();
                                checkBox3.Checked = true;

                            }
                        }
                    }
                }
                else
                {
                    Güncelle();
                    checkBox3.Checked = true;
                }
            }
            catch (Exception)
            {

                throw;
            }

        }


        private void temizle()
        {
            textBox11.Text = "";
            textBox15.Text = "";
            textBox14.Text = "";
            comboBox7.SelectedIndex = -1;
            comboBox5.SelectedIndex = -1;
            textBox2.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox12.Text = "";
            textBox17.Text = "";
        }
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (checkBox1.Checked == true)
                {
                    veridoldur();
                    checkBox3.Checked = false;
                }
                if (checkBox1.Checked == false)
                {
                    temizle();
                    checkBox3.Checked = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Güncelle Modu : " + ex.Message);
            }

        }

        private void pictureBox8_Click(object sender, EventArgs e)
        {


        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (textBox12.Text != "0")
                {
                    toplamfiyat = Convert.ToDouble(textBox15.Text) * Convert.ToDouble(textBox12.Text);
                    textBox17.Text = toplamfiyat.ToString("0.##");
                }
            }
            catch (Exception)
            {

            }

        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (comboBox1.Text != "")
            //{
            resimgetir();
            //fiyatbilgisigetir();
            //hgfiyatgetir();
            comboBox6.SelectedIndex = -1;
            comboBox7.SelectedIndex = -1;
        }
        //else
        //{
        //    MessageBox.Show("Lütfen önce müşteri seçiniz!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //}

        double renkstok;
        private void dataGridView3_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (dataGridView3.Rows[e.RowIndex].Cells["SiparisNo"].Value?.ToString() == "Toplam:")
            {
                dataGridView3.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Red;
                dataGridView3.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.White;
            }
        }
        void AyniSipGt()
        {
            if (comboBox1.Text != "" && comboBox6.Text != "")
            {
                string kayit = @"SELECT 
                    SiparisNo, 
                    SUM(ISNULL(TRY_CAST(REPLACE(ToplamM2, ',', '.') AS decimal(12, 2)), 0)) AS M2,
                    SiparişTarihi,
                    Renk,
                    Müşteri
                FROM Siparişler
                WHERE Müşteri=@Müşteri and Renk = @Renk
                GROUP BY Müşteri, Renk, SiparisNo, SiparişTarihi";

                string kayit2 = @"SELECT 
                    SIPARISNO AS SiparisNo, 
                    SUM(ISNULL(TRY_CAST(REPLACE(M2, ',', '.') AS decimal(12, 2)), 0)) AS M2,
                    SIPARISTARIHI AS SiparişTarihi,
                    RENK,
                    MUSTERI AS Müşteri
                FROM [Modeks_Eski].[dbo].[SIPARISLER] 
                WHERE MUSTERI = @Müşteri and Renk = @Renk
                GROUP BY MUSTERI, RENK, SIPARISNO, SIPARISTARIHI";

                SqlCommand command1 = new SqlCommand(kayit, bgl.baglanti());
                command1.Parameters.AddWithValue("@Müşteri", comboBox1.Text);
                command1.Parameters.AddWithValue("@Renk", comboBox6.Text);
                SqlDataAdapter adapter1 = new SqlDataAdapter(command1);
                System.Data.DataTable dt1 = new System.Data.DataTable();
                adapter1.Fill(dt1);

                SqlCommand command2 = new SqlCommand(kayit2, bgl.baglanti());
                command2.Parameters.AddWithValue("@Müşteri", comboBox1.Text);
                command2.Parameters.AddWithValue("@Renk", comboBox6.Text);
                SqlDataAdapter adapter2 = new SqlDataAdapter(command2);
                System.Data.DataTable dt2 = new System.Data.DataTable();
                adapter2.Fill(dt2);

                System.Data.DataTable mergedTable = dt1.Clone(); // kolon yapısını kopyala
                foreach (DataRow row in dt1.Rows)
                {
                    mergedTable.ImportRow(row);
                }
                foreach (DataRow row in dt2.Rows)
                {
                    mergedTable.ImportRow(row);
                }

                dataGridView3.RowPrePaint += new DataGridViewRowPrePaintEventHandler(dataGridView3_RowPrePaint);
                dataGridView3.DataSource = mergedTable;


            }

        }
        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            AyniSipGt();
            resimrenkgetir();
            try
            {
                renkstok = 0;
                fiyatbilgisigetir();
                hgfiyatgetir();
                RSMfiyatgetir();
                SOFTfiyatgetir();
                ULTRASOFTfiyatgetir();

                comboBox7.SelectedIndex = -1;
                StokGetir();
                if (textBox2.Text != "" && textBox6.Text != "" && textBox7.Text != "" && textBox2.Text != "0" && textBox6.Text != "0" && textBox7.Text != "0")
                {
                    formül();
                    toplamfiyat = Convert.ToDouble(textBox15.Text) * Convert.ToDouble(textBox12.Text);
                    textBox17.Text = toplamfiyat.ToString("0.##");
                }

                renkstok = 0;
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT * FROM Stoklar where Malzeme=@Malzeme";
                komut.Parameters.AddWithValue("@Malzeme", comboBox6.Text);
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    renkstok = Convert.ToDouble(dr["Kalan"].ToString());
                }

                label71.Text = comboBox6.Text + " = " + renkstok.ToString("0.##") + " m²";
            }
            catch (Exception)
            {

                throw;
            }




        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            özellikfiyatekle();


            double value1 = string.IsNullOrWhiteSpace(textBox15.Text) ? 0 : Convert.ToDouble(textBox15.Text);
            double value2 = string.IsNullOrWhiteSpace(textBox12.Text) ? 0 : Convert.ToDouble(textBox12.Text);

            double toplamfiyat = value1 * value2; textBox17.Text = toplamfiyat.ToString();
        }
        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            if (textBox14.Text == "")
            {
                textBox14.Text = "0";
            }
        }

        private void textBox14_Click(object sender, EventArgs e)
        {
            textBox14.Text = "";
        }

        private void comboBox4_Click(object sender, EventArgs e)
        {
            //if (comboBox1.Text == "")
            //{
            //    MessageBox.Show("Lütfen önce müşteri seçiniz!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
        }

        double mevcutstok;
        double kullanılanstok;
        double mevcutstokmdf;
        double kullanılanstokmdf;
        double mevcutstoktutkal;
        double kullanılanstoktutkal;
        double harcananstok;
        double harcananstokmdf;
        double harcananstoktutkal;
        private void button2_Click(object sender, EventArgs e)
        {
            //string sorgu = "UPDATE Siparişler SET Onay=@Onay WHERE SiparisNo=@SiparisNo";
            //SqlCommand komut;
            //komut = new SqlCommand(sorgu, bgl.baglanti());
            //komut.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
            //komut.Parameters.AddWithValue("@Onay", "Onaylandı");
            //komut.ExecuteNonQuery();
            //MessageBox.Show("Sipariş onaylanmıştır.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //liste();

            //for (int i = 0; i < dataGridView1.RowCount-1; i++)
            //{
            //    //m2
            //    SqlCommand komut2 = new SqlCommand();
            //    komut2.CommandText = "SELECT * FROM Stoklar where Malzeme=@Malzeme";
            //    komut2.Parameters.AddWithValue("@Malzeme", dataGridView1.Rows[i].Cells["Renk"].Value.ToString());
            //    komut2.Connection = bgl.baglanti();
            //    komut2.CommandType = CommandType.Text;

            //    SqlDataReader dr;
            //    dr = komut2.ExecuteReader();
            //    while (dr.Read())
            //    {
            //        mevcutstok = Convert.ToDouble(dr["MevcutStok"].ToString());
            //    }

            //    harcananstok = Convert.ToDouble(dataGridView1.Rows[i].Cells["M2"].Value.ToString()) * 2;
            //    string sorgu2 = "UPDATE Stoklar SET MevcutStok=@p1 WHERE Malzeme=@Malzeme";
            //    SqlCommand cmd2;
            //    cmd2 = new SqlCommand(sorgu2, bgl.baglanti());
            //    cmd2.Parameters.AddWithValue("@Malzeme", dataGridView1.Rows[i].Cells["Renk"].Value.ToString());
            //    cmd2.Parameters.AddWithValue("@p1", (mevcutstok - harcananstok).ToString());
            //    cmd2.ExecuteNonQuery();

            //    //----------------------------------------------------------------------------------------------//mdf
            //    SqlCommand komut3 = new SqlCommand();
            //    komut3.CommandText = "SELECT * FROM Stoklar where Malzeme=@Malzeme";
            //    komut3.Parameters.AddWithValue("@Malzeme", "18 MM TEKYÜZ MDF");
            //    komut3.Connection = bgl.baglanti();
            //    komut3.CommandType = CommandType.Text;

            //    SqlDataReader dr3;
            //    dr3 = komut3.ExecuteReader();
            //    while (dr3.Read())
            //    {
            //        mevcutstokmdf = Convert.ToDouble(dr3["MevcutStok"].ToString());
            //    }

            //    harcananstokmdf = Convert.ToDouble(dataGridView1.Rows[i].Cells["M2"].Value.ToString()) / 20;
            //    string sorgu3 = "UPDATE Stoklar SET MevcutStok=@p1 WHERE Malzeme=@Malzeme";
            //    SqlCommand cmd3;
            //    cmd3 = new SqlCommand(sorgu3, bgl.baglanti());
            //    cmd3.Parameters.AddWithValue("@Malzeme", "18 MM TEKYÜZ MDF");
            //    cmd3.Parameters.AddWithValue("@p1", (mevcutstokmdf - harcananstokmdf).ToString());
            //    cmd3.ExecuteNonQuery();


            //    //----------------------------------------------------------------------------------------------//tutkal
            //    SqlCommand komut4 = new SqlCommand();
            //    komut4.CommandText = "SELECT * FROM Stoklar where Malzeme=@Malzeme";
            //    komut4.Parameters.AddWithValue("@Malzeme", "TUTKAL");
            //    komut4.Connection = bgl.baglanti();
            //    komut4.CommandType = CommandType.Text;

            //    SqlDataReader dr4;
            //    dr4 = komut4.ExecuteReader();
            //    while (dr4.Read())
            //    {
            //        mevcutstoktutkal = Convert.ToDouble(dr4["MevcutStok"].ToString());
            //    }

            //    harcananstoktutkal = Convert.ToDouble(dataGridView1.Rows[i].Cells["M2"].Value.ToString()) / 4;
            //    string sorgu4 = "UPDATE Stoklar SET MevcutStok=@p1 WHERE Malzeme=@Malzeme";
            //    SqlCommand cmd4;
            //    cmd4 = new SqlCommand(sorgu4, bgl.baglanti());
            //    cmd4.Parameters.AddWithValue("@Malzeme", "TUTKAL");
            //    cmd4.Parameters.AddWithValue("@p1", (mevcutstoktutkal - harcananstoktutkal).ToString());
            //    cmd4.ExecuteNonQuery();

            //    StokGetir();
            //    TutkalStok();
            //    MDFStok();
            //}

            //checkBox1.Visible = false;
            //button9.Visible = false;
            //button10.Visible = false;
            //button8.Visible = false;
            //button1.Enabled = false;
            //button2.Enabled = false;
            //button3.Enabled = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }
        bool iskontoYap = false;
        private void İskontoGüncelle()
        {
            try
            {
                if (Convert.ToDouble(textBox16.Text) != 0 && textBox16.Text != "" && textBox16.Text != null && Convert.ToDouble(label63.Text) != 0)
                {
                    iskonto = ((Convert.ToDouble(label46.Text)) * Convert.ToDouble(textBox16.Text)) / 100;
                    label42.Text = iskonto.ToString("0.##");
                    Toplam();
                    AraToplam();
                    string sorgu = "UPDATE Siparişler SET İskonto=@İskonto,İskontoOrani=@İskontoOrani where SiparisNo=@p1 AND AnaSiparişMi=@p2";
                    SqlCommand komut;
                    komut = new SqlCommand(sorgu, bgl.baglanti());
                    komut.Parameters.AddWithValue("@p1", textBox1.Text);
                    komut.Parameters.AddWithValue("@p2", "Evet");
                    komut.Parameters.AddWithValue("@İskonto", label42.Text);
                    komut.Parameters.AddWithValue("@İskontoOrani", textBox16.Text);
                    komut.ExecuteNonQuery();
                    DDSHesaplama();
                    GenelToplam();
                }
                else if (Convert.ToDouble(textBox16.Text) == 0 && Convert.ToDouble(label63.Text) > 0 && textBox16.Text != "" && textBox16.Text != null)
                {
                    if (iskontoYap)
                    {
                        iskonto = (Convert.ToDouble(label46.Text) * Convert.ToDouble(textBox16.Text)) / 100;
                        label42.Text = iskonto.ToString("0.##");
                        Toplam();
                        AraToplam();
                        string sorgu = "UPDATE Siparişler SET İskonto=@İskonto,İskontoOrani=@İskontoOrani where SiparisNo=@p1 AND AnaSiparişMi=@p2";
                        SqlCommand komut;
                        komut = new SqlCommand(sorgu, bgl.baglanti());
                        komut.Parameters.AddWithValue("@p1", textBox1.Text);
                        komut.Parameters.AddWithValue("@p2", "Evet");
                        komut.Parameters.AddWithValue("@İskonto", label42.Text);
                        komut.Parameters.AddWithValue("@İskontoOrani", textBox16.Text);
                        komut.ExecuteNonQuery();
                        DDSHesaplama();
                        GenelToplam();
                    }
                }
            }
            catch (Exception)
            {

            }
        }
        private void button11_Click_1(object sender, EventArgs e)
        {
            iskontoYap = true;
            if (textBox16.Text == "")
            {
                textBox16.Text = "0";
            }
            İskontoGüncelle();

        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                groupBox5.Visible = true;
            }
            else
            {
                groupBox5.Visible = false;
            }
        }

        string sipTipi, sevkTürü;
        private void button1_Click(object sender, EventArgs e)
        {
            //if (comboBox1.Text != "" && comboBox2.Text != "" && comboBox3.Text != "")
            //{
            try
            {
                OpenFileDialog file = new OpenFileDialog();
                file.Filter = "Excel Dosyası |*.xls; *.xlsx";
                file.ShowDialog();
                string tamYol = file.FileName;
                string baglantiAdresi = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + tamYol + ";Extended Properties='Excel 12.0;IMEX=1;'";

                Excel.Application excel = new Excel.Application();
                excel.Visible = false;
                object Missing = Type.Missing;
                Excel.Workbook workbook = excel.Workbooks.Open(tamYol);
                Worksheet sheet1 = (Worksheet)workbook.Sheets["SABLON"];
                if (sheet1.Cells[6, 2].Value != null && sheet1.Cells[6, 2].Value.ToString() != "")
                {
                    textBox18.Text = sheet1.Cells[6, 2].Value.ToString();
                    sipTipi = sheet1.Cells[8, 4].Value.ToString();
                    sevkTürü = sheet1.Cells[7, 4].Value.ToString();
                }
                OleDbConnection baglanti = new OleDbConnection(baglantiAdresi);
                OleDbCommand komut = new OleDbCommand("Select * From [SABLON$A10:G50]", baglanti);
                OleDbDataAdapter da = new OleDbDataAdapter(komut);
                System.Data.DataTable data = new System.Data.DataTable();
                da.Fill(data);
                dataGridView2.DataSource = data;
                baglanti.Close();
                bool eslesme = false;

                foreach (var item in comboBox1.Items)
                {
                    if (item.ToString().ToLower() == textBox18.Text.ToLower())
                    {
                        comboBox1.Text = textBox18.Text;
                        eslesme = true;
                        for (int i = 0; i < dataGridView2.RowCount - 1; i++)
                        {
                            if (dataGridView2.Rows[i].Cells[0].Value.ToString() != "")
                            {
                                string kayit = "insert into Siparişler(SiparisNo,Onay,Müşteri,SiparişTipi,SevkTürü,BID,SiparişTarihi,TeslimTarihi,Nott,Model,Renk,Özellik,Boy,En,Adet,M2,TasarımÜcreti,M2Fiyat,Fiyat,Fiyat2,AnaSiparişMi,Adres,Telefon,Aşama,EkleyenKullanici) values (@SiparisNo,@Onay,@Müşteri,@SiparişTipi,@SevkTürü,@BID,@SiparişTarihi,@TeslimTarihi,@Nott,@Model,@Renk,@Özellik,@Boy,@En,@Adet,@M2,@TasarımÜcreti,@M2Fiyat,@Fiyat,@Fiyat2,@AnaSiparişMi,@Adres,@Telefon,@Aşama, @EkleyenKullanici)";
                                SqlCommand komut2 = new SqlCommand(kayit, bgl.baglanti());
                                komut2.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
                                komut2.Parameters.AddWithValue("@Onay", "Onay Bekliyor");
                                komut2.Parameters.AddWithValue("@Aşama", "Onay Bekliyor");
                                komut2.Parameters.AddWithValue("@Müşteri", comboBox1.Text);
                                komut2.Parameters.AddWithValue("@SiparişTipi", comboBox2.Text);
                                komut2.Parameters.AddWithValue("@SevkTürü", comboBox3.Text);
                                komut2.Parameters.AddWithValue("@BID", textBox8.Text);
                                komut2.Parameters.AddWithValue("@SiparişTarihi", dateTimePicker1.Value);
                                komut2.Parameters.AddWithValue("@TeslimTarihi", textBox10.Text);
                                komut2.Parameters.AddWithValue("@Nott", textBox11.Text);
                                komut2.Parameters.AddWithValue("@Model", dataGridView2.Rows[i].Cells[0].Value);
                                komut2.Parameters.AddWithValue("@Renk", dataGridView2.Rows[i].Cells[1].Value);
                                komut2.Parameters.AddWithValue("@EkleyenKullanici", kullaniciadi);
                                comboBox4.SelectedItem = dataGridView2.Rows[i].Cells[0].Value;
                                comboBox6.SelectedItem = dataGridView2.Rows[i].Cells[1].Value;

                                if (sipTipi == "Спешно")
                                    comboBox2.SelectedItem = "Acil";
                                else
                                    comboBox2.SelectedItem = "Normal";

                                if (sevkTürü == "КАРГО")
                                    comboBox3.SelectedItem = "Kargo";
                                else
                                    comboBox3.SelectedItem = "Fabrika";

                                textBox2.Text = dataGridView2.Rows[i].Cells[3].Value.ToString();
                                textBox6.Text = dataGridView2.Rows[i].Cells[4].Value.ToString();
                                textBox7.Text = dataGridView2.Rows[i].Cells[5].Value.ToString();
                                textBox12.Text = dataGridView2.Rows[i].Cells[6].Value.ToString();
                                if (dataGridView2.Rows[i].Cells[2].Value.ToString() == "1-Cekmece butun")
                                {
                                    dataGridView2.Rows[i].Cells[2].Value = "CEKMECE BUTUN";
                                    comboBox7.Text = "CEKMECE BUTUN";

                                }
                                else if (dataGridView2.Rows[i].Cells[2].Value.ToString() == "2-Cızım")
                                {
                                    dataGridView2.Rows[i].Cells[2].Value = "CIZIM";
                                    comboBox7.Text = "CIZIM";
                                }
                                else if (dataGridView2.Rows[i].Cells[2].Value.ToString() == "3-Cam")
                                {
                                    dataGridView2.Rows[i].Cells[2].Value = "Cam";
                                    comboBox7.Text = "Cam";
                                }
                                else if (dataGridView2.Rows[i].Cells[2].Value.ToString() == "4-Cam 2 Kafes")
                                {
                                    dataGridView2.Rows[i].Cells[2].Value = "Cam 2 Kafes";
                                    comboBox7.Text = "Cam 2 Kafes";
                                }
                                else if (dataGridView2.Rows[i].Cells[2].Value.ToString() == "5-Cam 4 Kafes")
                                {
                                    dataGridView2.Rows[i].Cells[2].Value = "Cam 4 Kafes";
                                    comboBox7.Text = "Cam 4 Kafes";
                                }
                                else if (dataGridView2.Rows[i].Cells[2].Value.ToString() == "6-Cam 6 Kafes")
                                {
                                    dataGridView2.Rows[i].Cells[2].Value = "Cam 6 Kafes";
                                    comboBox7.Text = "Cam 6 Kafes";
                                }
                                else if (dataGridView2.Rows[i].Cells[2].Value.ToString() == "7-Cam 8 Kafes")
                                {
                                    dataGridView2.Rows[i].Cells[2].Value = "Cam 8 Kafes";
                                    comboBox7.Text = "Cam 8 Kafes";
                                }
                                else if (dataGridView2.Rows[i].Cells[2].Value.ToString() == "8-Tac T1")
                                {
                                    dataGridView2.Rows[i].Cells[2].Value = "T1";
                                    comboBox7.Text = "T1";
                                }
                                else if (dataGridView2.Rows[i].Cells[2].Value.ToString() == "9-Tac T2")
                                {
                                    dataGridView2.Rows[i].Cells[2].Value = "T2";
                                    comboBox7.Text = "T2";
                                }
                                else
                                {
                                    comboBox7.Text = "CEKMECE BUTUN";
                                }
                                özellikfiyatekle();
                                komut2.Parameters.AddWithValue("@Özellik", dataGridView2.Rows[i].Cells[2].Value);
                                komut2.Parameters.AddWithValue("@Boy", (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value)).ToString("0.##"));
                                komut2.Parameters.AddWithValue("@En", (Convert.ToDouble(dataGridView2.Rows[i].Cells[4].Value)).ToString("0.##"));
                                komut2.Parameters.AddWithValue("@Adet", (Convert.ToDouble(dataGridView2.Rows[i].Cells[5].Value)).ToString("0.##"));
                                komut2.Parameters.AddWithValue("@M2", (Convert.ToDouble(dataGridView2.Rows[i].Cells[6].Value)).ToString("0.##"));
                                komut2.Parameters.AddWithValue("@TasarımÜcreti", "0");
                                komut2.Parameters.AddWithValue("@M2Fiyat", (Convert.ToDouble(textBox15.Text)).ToString("0.##"));
                                komut2.Parameters.AddWithValue("@Fiyat", (Convert.ToDouble(textBox15.Text)).ToString("0.##"));
                                komut2.Parameters.AddWithValue("@Fiyat2", (Convert.ToDouble(textBox15.Text) * Convert.ToDouble(textBox12.Text)).ToString("0.##"));
                                //if (i == 0)
                                //{
                                komut2.Parameters.AddWithValue("@AnaSiparişMi", "Evet");
                                //}
                                //else
                                //{
                                //komut2.Parameters.AddWithValue("@AnaSiparişMi", " ");

                                //}
                                komut2.Parameters.AddWithValue("@Adres", textBox3.Text);
                                komut2.Parameters.AddWithValue("@Telefon", textBox4.Text);
                                komut2.ExecuteNonQuery();

                                İstatistik();
                                bgl.baglanti().Close();
                                SqlConnection.ClearPool(bgl.baglanti());
                            }
                        }
                        comboBox4.SelectedIndex = -1;
                        comboBox6.SelectedIndex = -1;
                        textBox15.Text = "0";
                        liste();
                        birlesikliste();
                        M2Toplat();
                        KapakAdetToplat();
                        TasarımUcretıToplat();
                        KapakToplam();
                        M2KapakSayısı();
                        if (comboBox2.Text == "Acil")
                        {
                            AcilHesaplama();
                        }
                        M2KapakFarkıHesaplama();
                        KargoHesaplama();
                        Toplam();
                        AraToplam();
                        DDSHesaplama();
                        GenelToplam();



                        string sorgu = "UPDATE Siparişler SET ToplamM2=@ToplamM2,ToplamAdet=@ToplamAdet,ToplamFiyat=@ToplamFiyat,ToplamTasarımÜcreti=@ToplamTasarımÜcreti,İskonto=@İskonto,Kargo=@Kargo,AcilFarkı=@AcilFarkı,M2KapakFarkı=@M2KapakFarkı,M2KapakAdet=@M2KapakAdet,DDS=@DDS WHERE SiparisNo=@SiparisNo AND AnaSiparişMi=@AnaSiparişMi";
                        SqlCommand komut3;
                        komut3 = new SqlCommand(sorgu, bgl.baglanti());
                        komut3.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
                        komut3.Parameters.AddWithValue("@AnaSiparişMi", "Evet");
                        komut3.Parameters.AddWithValue("@ToplamM2", label27.Text);
                        komut3.Parameters.AddWithValue("@ToplamAdet", label28.Text);
                        komut3.Parameters.AddWithValue("@ToplamFiyat", label63.Text);
                        komut3.Parameters.AddWithValue("@ToplamTasarımÜcreti", label34.Text);
                        komut3.Parameters.AddWithValue("@İskonto", textBox16.Text);
                        komut3.Parameters.AddWithValue("@Kargo", label40.Text);
                        komut3.Parameters.AddWithValue("@AcilFarkı", label38.Text);
                        komut3.Parameters.AddWithValue("@M2KapakFarkı", label36.Text);
                        komut3.Parameters.AddWithValue("@M2KapakAdet", label30.Text);
                        komut3.Parameters.AddWithValue("@DDS", label44.Text);

                        komut3.ExecuteNonQuery();
                        bgl.baglanti().Close();
                        SqlConnection.ClearPool(bgl.baglanti());
                        break;
                    }
                }

                if (!eslesme)
                {

                    DialogResult dialogResult_sablon = MessageBox.Show("Şablondaki müşteri bulunamadı. Eklemek istiyor musun?", "Hata", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (dialogResult_sablon == DialogResult.Yes)
                    {
                        for (int i = 0; i < dataGridView2.RowCount - 1; i++)
                        {
                            if (dataGridView2.Rows[i].Cells[0].Value.ToString() != "")
                            {
                                string kayit = "insert into Siparişler(SiparisNo,Onay,Müşteri,SiparişTipi,SevkTürü,BID,SiparişTarihi,TeslimTarihi,Nott,Model,Renk,Özellik,Boy,En,Adet,M2,TasarımÜcreti,M2Fiyat,Fiyat2,AnaSiparişMi,Adres,Telefon,Aşama,EkleyenKullanici) values (@SiparisNo,@Onay,@Müşteri,@SiparişTipi,@SevkTürü,@BID,@SiparişTarihi,@TeslimTarihi,@Nott,@Model,@Renk,@Özellik,@Boy,@En,@Adet,@M2,@TasarımÜcreti,@M2Fiyat,@Fiyat2,@AnaSiparişMi,@Adres,@Telefon,@Aşama,@EkleyenKullanici)";
                                SqlCommand komut2 = new SqlCommand(kayit, bgl.baglanti());
                                komut2.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
                                komut2.Parameters.AddWithValue("@Onay", "Onay Bekliyor");
                                komut2.Parameters.AddWithValue("@Aşama", "Onay Bekliyor");
                                komut2.Parameters.AddWithValue("@Müşteri", comboBox1.Text);
                                komut2.Parameters.AddWithValue("@SiparişTipi", comboBox2.Text);
                                komut2.Parameters.AddWithValue("@SevkTürü", comboBox3.Text);
                                komut2.Parameters.AddWithValue("@BID", textBox8.Text);
                                komut2.Parameters.AddWithValue("@SiparişTarihi", dateTimePicker1.Value);
                                komut2.Parameters.AddWithValue("@TeslimTarihi", textBox10.Text);
                                komut2.Parameters.AddWithValue("@Nott", textBox11.Text);
                                komut2.Parameters.AddWithValue("@Model", dataGridView2.Rows[i].Cells[0].Value);
                                komut2.Parameters.AddWithValue("@Renk", dataGridView2.Rows[i].Cells[1].Value);
                                komut2.Parameters.AddWithValue("@EkleyenKullanici", kullaniciadi);

                                comboBox4.SelectedItem = dataGridView2.Rows[i].Cells[0].Value;
                                comboBox6.SelectedItem = dataGridView2.Rows[i].Cells[1].Value;

                                if (sipTipi == "Спешно")
                                    comboBox2.SelectedItem = "Acil";
                                else
                                    comboBox2.SelectedItem = "Normal";

                                if (sevkTürü == "КАРГО")
                                    comboBox3.SelectedItem = "Kargo";
                                else
                                    comboBox3.SelectedItem = "Fabrika";

                                textBox2.Text = dataGridView2.Rows[i].Cells[3].Value.ToString();
                                textBox6.Text = dataGridView2.Rows[i].Cells[4].Value.ToString();
                                textBox7.Text = dataGridView2.Rows[i].Cells[5].Value.ToString();
                                textBox12.Text = dataGridView2.Rows[i].Cells[6].Value.ToString();
                                if (dataGridView2.Rows[i].Cells[2].Value.ToString() == "1-Cekmece butun")
                                {
                                    dataGridView2.Rows[i].Cells[2].Value = "CEKMECE BUTUN";
                                    comboBox7.Text = "CEKMECE BUTUN";

                                }
                                else if (dataGridView2.Rows[i].Cells[2].Value.ToString() == "2-Cızım")
                                {
                                    dataGridView2.Rows[i].Cells[2].Value = "CIZIM";
                                    comboBox7.Text = "CIZIM";
                                }
                                else if (dataGridView2.Rows[i].Cells[2].Value.ToString() == "3-Cam")
                                {
                                    dataGridView2.Rows[i].Cells[2].Value = "Cam";
                                    comboBox7.Text = "Cam";
                                }
                                else if (dataGridView2.Rows[i].Cells[2].Value.ToString() == "4-Cam 2 Kafes")
                                {
                                    dataGridView2.Rows[i].Cells[2].Value = "Cam 2 Kafes";
                                    comboBox7.Text = "Cam 2 Kafes";
                                }
                                else if (dataGridView2.Rows[i].Cells[2].Value.ToString() == "5-Cam 4 Kafes")
                                {
                                    dataGridView2.Rows[i].Cells[2].Value = "Cam 4 Kafes";
                                    comboBox7.Text = "Cam 4 Kafes";
                                }
                                else if (dataGridView2.Rows[i].Cells[2].Value.ToString() == "6-Cam 6 Kafes")
                                {
                                    dataGridView2.Rows[i].Cells[2].Value = "Cam 6 Kafes";
                                    comboBox7.Text = "Cam 6 Kafes";
                                }
                                else if (dataGridView2.Rows[i].Cells[2].Value.ToString() == "7-Cam 8 Kafes")
                                {
                                    dataGridView2.Rows[i].Cells[2].Value = "Cam 8 Kafes";
                                    comboBox7.Text = "Cam 8 Kafes";
                                }
                                else if (dataGridView2.Rows[i].Cells[2].Value.ToString() == "8-Tac T1")
                                {
                                    dataGridView2.Rows[i].Cells[2].Value = "T1";
                                    comboBox7.Text = "T1";
                                }
                                else if (dataGridView2.Rows[i].Cells[2].Value.ToString() == "9-Tac T2")
                                {
                                    dataGridView2.Rows[i].Cells[2].Value = "T2";
                                    comboBox7.Text = "T2";
                                }
                                else
                                {
                                    comboBox7.Text = "CEKMECE BUTUN";
                                }
                                özellikfiyatekle();
                                komut2.Parameters.AddWithValue("@Özellik", dataGridView2.Rows[i].Cells[2].Value);
                                komut2.Parameters.AddWithValue("@Boy", (Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value)).ToString("0.##"));
                                komut2.Parameters.AddWithValue("@En", (Convert.ToDouble(dataGridView2.Rows[i].Cells[4].Value)).ToString("0.##"));
                                komut2.Parameters.AddWithValue("@Adet", (Convert.ToDouble(dataGridView2.Rows[i].Cells[5].Value)).ToString("0.##"));
                                komut2.Parameters.AddWithValue("@M2", (Convert.ToDouble(dataGridView2.Rows[i].Cells[6].Value)).ToString("0.##"));
                                komut2.Parameters.AddWithValue("@TasarımÜcreti", "0");
                                komut2.Parameters.AddWithValue("@M2Fiyat", (Convert.ToDouble(textBox15.Text)).ToString("0.##"));
                                komut2.Parameters.AddWithValue("@Fiyat2", (Convert.ToDouble(textBox15.Text) * Convert.ToDouble(textBox12.Text)).ToString("0.##"));
                                //if (i == 0)
                                //{
                                komut2.Parameters.AddWithValue("@AnaSiparişMi", "Evet");
                                //}
                                //else
                                //{
                                //komut2.Parameters.AddWithValue("@AnaSiparişMi", " ");

                                //}
                                komut2.Parameters.AddWithValue("@Adres", textBox3.Text);
                                komut2.Parameters.AddWithValue("@Telefon", textBox4.Text);
                                komut2.ExecuteNonQuery();

                                İstatistik();
                                bgl.baglanti().Close();
                                SqlConnection.ClearPool(bgl.baglanti());
                            }
                        }
                        comboBox4.SelectedIndex = -1;
                        comboBox6.SelectedIndex = -1;
                        textBox15.Text = "0";
                        liste();
                        birlesikliste();
                        M2Toplat();
                        KapakAdetToplat();
                        TasarımUcretıToplat();
                        KapakToplam();
                        M2KapakSayısı();
                        if (comboBox2.Text == "Acil")
                        {
                            AcilHesaplama();
                        }
                        M2KapakFarkıHesaplama();
                        KargoHesaplama();
                        Toplam();
                        AraToplam();
                        DDSHesaplama();
                        GenelToplam();

                        string sorgu = "UPDATE Siparişler SET ToplamM2=@ToplamM2,ToplamAdet=@ToplamAdet,ToplamFiyat=@ToplamFiyat,ToplamTasarımÜcreti=@ToplamTasarımÜcreti,İskonto=@İskonto,Kargo=@Kargo,AcilFarkı=@AcilFarkı,M2KapakFarkı=@M2KapakFarkı,M2KapakAdet=@M2KapakAdet,DDS=@DDS WHERE SiparisNo=@SiparisNo AND AnaSiparişMi=@AnaSiparişMi";
                        SqlCommand komut3;
                        komut3 = new SqlCommand(sorgu, bgl.baglanti());
                        komut3.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
                        komut3.Parameters.AddWithValue("@AnaSiparişMi", "Evet");
                        komut3.Parameters.AddWithValue("@ToplamM2", label27.Text);
                        komut3.Parameters.AddWithValue("@ToplamAdet", label28.Text);
                        komut3.Parameters.AddWithValue("@ToplamFiyat", label63.Text);
                        komut3.Parameters.AddWithValue("@ToplamTasarımÜcreti", label34.Text);
                        komut3.Parameters.AddWithValue("@İskonto", textBox16.Text);
                        komut3.Parameters.AddWithValue("@Kargo", label40.Text);
                        komut3.Parameters.AddWithValue("@AcilFarkı", label38.Text);
                        komut3.Parameters.AddWithValue("@M2KapakFarkı", label36.Text);
                        komut3.Parameters.AddWithValue("@M2KapakAdet", label30.Text);
                        komut3.Parameters.AddWithValue("@DDS", label44.Text);
                        komut3.ExecuteNonQuery();
                        bgl.baglanti().Close();
                        SqlConnection.ClearPool(bgl.baglanti());
                    }
                    else
                    {

                    }
                }
                SiparişBölmeSayısı();
                if (bölmekayitsayisi > 1)
                {
                    MessageBox.Show("İskonto düzenlemeniz pasif hale gelmiştir, önce siparişi bölünüz..");
                    checkBox2.Enabled = false;
                }
                else
                    checkBox2.Enabled = true;
                workbook.Close();
                excel.Quit();
                foreach (var process in Process.GetProcessesByName("EXCEL"))
                {
                    process.Kill();
                }
                // Nesneleri serbest bırakma
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet1);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

                // Bellekten nesneleri temizleme
                sheet1 = null;
                workbook = null;
                excel = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();

            }
            catch (Exception)
            {

            }
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void CSV()
        {
            try
            {
                Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                if (xlApp == null)
                {
                    MessageBox.Show("Sisteminizde Excel kurulu değil...");
                    return;
                }

                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    if (dataGridView1.Rows[i].Cells[4].Value.ToString().Length > 3)
                    {
                        // boy 20 den küçükse
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString()) < 20)
                        {
                            string cellValue = dataGridView1.Rows[i].Cells[1].Value.ToString();
                            cellValue = cellValue.Replace("-", " ");
                            xlWorkSheet.Cells[i + 1, 1] = cellValue + " MOON CEK" + dataGridView1.Rows[i].Cells[4].Value.ToString() + "," + Convert.ToString(Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString()) * 10) + "," + Convert.ToString(Convert.ToDouble(dataGridView1.Rows[i].Cells[6].Value.ToString()) * 10) + "," + dataGridView1.Rows[i].Cells[7].Value.ToString() + "," + textBox1.Text + "," + " 0" + (i + 1).ToString();
                        }
                        //değilse
                        else
                        {
                            string cellValue = dataGridView1.Rows[i].Cells[1].Value.ToString();
                            cellValue = cellValue.Replace("-", " ");
                            xlWorkSheet.Cells[i + 1, 1] = cellValue + " MOON " + dataGridView1.Rows[i].Cells[4].Value.ToString() + "," + Convert.ToString(Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString()) * 10) + "," + Convert.ToString(Convert.ToDouble(dataGridView1.Rows[i].Cells[6].Value.ToString()) * 10) + "," + dataGridView1.Rows[i].Cells[7].Value.ToString() + "," + textBox1.Text + "," + " 0" + (i + 1).ToString();
                        }
                    }
                    else
                    {
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString()) < 20)
                        {
                            string cellValue = dataGridView1.Rows[i].Cells[1].Value.ToString();
                            cellValue = cellValue.Replace("-", " ");
                            xlWorkSheet.Cells[i + 1, 1] = cellValue + " MOON CEK" + "," + Convert.ToString(Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString()) * 10) + "," + Convert.ToString(Convert.ToDouble(dataGridView1.Rows[i].Cells[6].Value.ToString()) * 10) + "," + dataGridView1.Rows[i].Cells[7].Value.ToString() + "," + textBox1.Text + "," + " 0" + (i + 1).ToString();
                        }
                        else
                        {
                            string cellValue = dataGridView1.Rows[i].Cells[1].Value.ToString();
                            cellValue = cellValue.Replace("-", " ");
                            xlWorkSheet.Cells[i + 1, 1] = cellValue + " MOON " + "," + Convert.ToString(Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value.ToString()) * 10) + "," + Convert.ToString(Convert.ToDouble(dataGridView1.Rows[i].Cells[6].Value.ToString()) * 10) + "," + dataGridView1.Rows[i].Cells[7].Value.ToString() + "," + textBox1.Text + "," + " 0" + (i + 1).ToString();
                        }
                    }
                }
                try
                {
                    string csvFilePath = "C:\\Modeks_Dosyalar\\csv\\" + textBox1.Text + ".csv";
                    xlWorkBook.SaveAs(csvFilePath, Excel.XlFileFormat.xlCSV, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();
                    foreach (var process in Process.GetProcessesByName("EXCEL"))
                    {
                        process.Kill();
                    }
                    // UTF-8 karakter kodlaması kullanarak dosyanın içeriğini güncelliyoruz.
                    string fileContent = System.IO.File.ReadAllText(csvFilePath);
                    fileContent = fileContent.Replace("\"", "");

                    // UTF-8 BOM olmadan dosyayı yazma
                    using (StreamWriter sw = new StreamWriter(csvFilePath, false, new UTF8Encoding(false)))
                    {
                        sw.Write(fileContent);
                    }

                    MessageBox.Show("Excel dosyası C:\\Modeks_Dosyalar\\csv\\" + textBox1.Text + ".csv adresinde oluşturuldu...");
                }
                catch (Exception)
                {
                    // Hata yönetimi burada yapılabilir.
                }
            }
            catch (Exception)
            {

                throw;
            }

        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            if (comboBox1.Text == "" || comboBox2.Text == "" || comboBox3.Text == "")
            {
                MessageBox.Show("Müşteri Kısmını, Sipariş Şeklini ve Sipariş Türünü Seçiniz!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {

                message = "";
                bölmesayısı = 0;
                if (dataGridView1.RowCount == 1)
                {
                    MessageBox.Show("Sipariş Yoktur!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    SiparişBölme();
                    if (siparisbölme == "böl")
                    {

                        for (int i = 0; i < bölmesayısı; i++)
                        {
                            string sorgu = "UPDATE Siparişler SET SiparisNo=@SiparisNo, Müşteri=@Müşteri, SiparişTipi=@SiparişTipi, SevkTürü=@SevkTürü,  Aşama=@Aşama, AnaSiparişMi=@AnaSiparişMi,ToplamTasarımÜcreti=@ToplamTasarımÜcreti,ToplamAdet=@ToplamAdet,ToplamFiyat=@ToplamFiyat,ToplamM2=@ToplamM2,İskonto=@İskonto,Kargo=@Kargo,AcilFarkı=@AcilFarkı,M2KapakFarkı=@M2KapakFarkı,DDS=@DDS, SiparişiYazan=@SiparişiYazan WHERE SiparisNo=@p2 AND Renk=@Renk";
                            SqlCommand komut;
                            komut = new SqlCommand(sorgu, bgl.baglanti());
                            komut.Parameters.AddWithValue("@p2", textBox1.Text);
                            komut.Parameters.AddWithValue("@Renk", comboBox8.Items[i].ToString());
                            komut.Parameters.AddWithValue("@SiparisNo", textBox1.Text + "-" + "" + (i + 1) + "");
                            komut.Parameters.AddWithValue("@Müşteri", comboBox1.Text);
                            komut.Parameters.AddWithValue("@SiparişTipi", comboBox2.Text);
                            komut.Parameters.AddWithValue("@SevkTürü", comboBox3.Text);
                            komut.Parameters.AddWithValue("@AnaSiparişMi", "Evet");
                            komut.Parameters.AddWithValue("@Aşama", "Onay Bekliyor");
                            komut.Parameters.AddWithValue("@ToplamTasarımÜcreti", dataGridView1.Rows[i].Cells[3].Value.ToString());
                            komut.Parameters.AddWithValue("@ToplamAdet", dataGridView1.Rows[i].Cells[7].Value.ToString());
                            kapaktoplam = Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString());
                            toplamM2 = Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString());
                            KargoHesaplama();
                            AcilHesaplama();
                            toplamKapakAdet = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString());
                            M2KapakFarkıHesaplama();
                            geneltoplam = (Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value.ToString()) + kargo + acil + m2kapakfarkı);
                            komut.Parameters.AddWithValue("@ToplamFiyat", geneltoplam.ToString());
                            komut.Parameters.AddWithValue("@ToplamM2", dataGridView1.Rows[i].Cells[8].Value.ToString());
                            if (label42.Text == null)
                                label42.Text = "0";
                            komut.Parameters.AddWithValue("@İskonto", label42.Text);
                            kapaktoplam = Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString());
                            toplamM2 = Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value.ToString());
                            KargoHesaplama();
                            geneltoplam = (Convert.ToDouble(dataGridView1.Rows[i].Cells[10].Value.ToString()) + Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value.ToString()) + kargo + acil + m2kapakfarkı);
                            AcilHesaplama();
                            toplamKapakAdet = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value.ToString());
                            M2KapakFarkıHesaplama();
                            DDSHesaplama();
                            komut.Parameters.AddWithValue("@Kargo", label40.Text);
                            komut.Parameters.AddWithValue("@AcilFarkı", label38.Text);
                            komut.Parameters.AddWithValue("@M2KapakFarkı", label36.Text);
                            komut.Parameters.AddWithValue("@DDS", "");
                            komut.Parameters.AddWithValue("@SiparişiYazan", label59.Text);
                            komut.ExecuteNonQuery();
                            bgl.baglanti().Close();
                            SqlConnection.ClearPool(bgl.baglanti());

                            SqlCommand verigetir = new SqlCommand();
                            verigetir.CommandText = "SELECT *FROM Siparişler WHERE SiparisNo=@SiparisNo";
                            verigetir.Parameters.AddWithValue("@SiparisNo", textBox1.Text + "-" + "" + (i + 1) + "");
                            verigetir.Connection = bgl.baglanti();
                            verigetir.CommandType = CommandType.Text;
                            SqlDataReader dr;
                            dr = verigetir.ExecuteReader();
                            toplamM2 = 0;
                            toplamKapakAdet = 0;
                            kapaktoplam = 0;
                            toplamfiyat = 0;
                            toplamTasarımUcreti = 0;
                            iskonto = 0;
                            while (dr.Read())
                            {

                                toplamM2 += Convert.ToDouble(dr["M2"]);
                                toplamKapakAdet += Convert.ToDouble(dr["Adet"]);
                                kapaktoplam += Convert.ToDouble(dr["Fiyat2"]);
                                toplamTasarımUcreti += Convert.ToDouble(dr["TasarımÜcreti"]);
                                iskonto += Convert.ToDouble(dr["İskonto"]);


                            }
                            KargoHesaplama();
                            AcilHesaplama();
                            M2KapakFarkıHesaplama();
                            toplamfiyat = kapaktoplam + kargo + acil + m2kapakfarkı + toplamTasarımUcreti;
                            geneltoplam = toplamfiyat;
                            DDSHesaplama();
                            bgl.baglanti().Close();
                            SqlConnection.ClearPool(bgl.baglanti());

                            dateTimePicker1.Text = DateTime.Now.ToString("dd.MM.yyyy HH:dd:ss");
                            if (comboBox2.Text == "Acil")
                            {
                                textBox10.Text = dateTimePicker1.Value.AddDays(5).ToString("dd.MM.yyyy HH:dd:ss");
                            }
                            else if (comboBox2.Text == "Normal")
                            {
                                textBox10.Text = dateTimePicker1.Value.AddDays(20).ToString("dd.MM.yyyy HH:dd:ss");
                            }
                            string sorgu2 = "UPDATE Siparişler SET Aşama='Onaylandı', ToplamM2=@ToplamM2,ToplamAdet=@ToplamAdet,ToplamFiyat=@ToplamFiyat,ToplamTasarımÜcreti=@ToplamTasarımÜcreti,İskonto=@İskonto,Kargo=@Kargo,AcilFarkı=@AcilFarkı,M2KapakFarkı=@M2KapakFarkı,M2KapakAdet=@M2KapakAdet,BID=@BID,Nott=@Nott,TeslimTarihi=@TeslimTarihi WHERE SiparisNo=@SiparisNo AND AnaSiparişMi=@AnaSiparişMi";
                            SqlCommand komut2;
                            komut2 = new SqlCommand(sorgu2, bgl.baglanti());
                            komut2.Parameters.AddWithValue("@SiparisNo", textBox1.Text + "-" + "" + (i + 1) + "");
                            komut2.Parameters.AddWithValue("@AnaSiparişMi", "Evet");
                            komut2.Parameters.AddWithValue("@ToplamM2", toplamM2.ToString());
                            komut2.Parameters.AddWithValue("@ToplamAdet", toplamKapakAdet.ToString());
                            komut2.Parameters.AddWithValue("@ToplamFiyat", toplamfiyat.ToString());
                            komut2.Parameters.AddWithValue("@ToplamTasarımÜcreti", toplamTasarımUcreti.ToString());
                            if (iskonto == null)
                                iskonto = 0;
                            komut2.Parameters.AddWithValue("@İskonto", iskonto.ToString());
                            komut2.Parameters.AddWithValue("@Kargo", kargo.ToString());
                            komut2.Parameters.AddWithValue("@AcilFarkı", acil.ToString());

                            komut2.Parameters.AddWithValue("@M2KapakFarkı", m2kapakfarkı.ToString());
                            komut2.Parameters.AddWithValue("@M2KapakAdet", m2kapaksayısı.ToString());
                            //komut2.Parameters.AddWithValue("@DDS", dds.ToString());
                            komut2.Parameters.AddWithValue("@BID", textBox8.Text);
                            komut2.Parameters.AddWithValue("@Nott", textBox11.Text);
                            //komut2.Parameters.AddWithValue("@SiparişTarihi", dateTimePicker1.Text);
                            komut2.Parameters.AddWithValue("@TeslimTarihi", textBox10.Text);
                            komut2.ExecuteNonQuery();
                            bgl.baglanti().Close();
                            SqlConnection.ClearPool(bgl.baglanti());
                        }
                        bgl.baglanti().Close();
                        SqlConnection.ClearPool(bgl.baglanti());
                        MessageBox.Show("Sipariş başarıyla bölünmüştür.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        liste();
                    }
                    else if (siparisbölme == "bölme")
                    {
                        Form3 frm2 = new Form3();
                        frm2.yetki = yetki;
                        frm2.kullaniciadi = kullaniciadi;
                        frm2.hangiformdan = hangiformdan;
                        this.Hide();
                        frm2.Show();
                    }
                    else
                    {
                        satıs_ıd_getir();
                        bsr_kullanici_getir();
                        BSR_Kaydet();
                        CSV();

                        KapakAdetToplat();
                        TasarımUcretıToplat();
                        KapakToplam();
                        M2KapakSayısı();
                        if (comboBox2.Text == "Acil")
                        {
                            AcilHesaplama();
                        }
                        M2KapakFarkıHesaplama();
                        KargoHesaplama();
                        Toplam();
                        AraToplam();
                        DDSHesaplama();
                        GenelToplam();

                        //dateTimePicker1.Text = DateTime.Now.ToString("dd.MM.yyyy HH:dd:ss");
                        if (comboBox2.Text == "Acil")
                        {
                            textBox10.Text = DateTime.Now.AddDays(5).ToString("dd.MM.yyyy HH:dd:ss");
                        }
                        else if (comboBox2.Text == "Normal")
                        {
                            textBox10.Text = DateTime.Now.AddDays(20).ToString("dd.MM.yyyy HH:dd:ss");
                        }
                        string sorgu = "UPDATE Siparişler SET ToplamFiyat=@ToplamFiyat, Aşama=@Aşama, Onay=@Onay, Müşteri=@Müşteri, SiparişTipi=@SiparişTipi, SevkTürü=@SevkTürü, SiparişiYazan=@SiparişiYazan, M2KapakAdet=@M2KapakAdet, OnayTarihi=@OnayTarihi,BID=@BID,Nott=@Nott,SiparişTarihi=@SiparişTarihi,TeslimTarihi=@TeslimTarihi, DDS=@DDS WHERE SiparisNo=@SiparisNo";
                        SqlCommand komut;
                        komut = new SqlCommand(sorgu, bgl.baglanti());
                        komut.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
                        komut.Parameters.AddWithValue("@ToplamFiyat", label63.Text);
                        komut.Parameters.AddWithValue("@Onay", "Onaylandı");
                        komut.Parameters.AddWithValue("@Aşama", "Onaylandı");
                        komut.Parameters.AddWithValue("@Müşteri", comboBox1.Text);
                        komut.Parameters.AddWithValue("@SiparişTipi", comboBox2.Text);
                        komut.Parameters.AddWithValue("@SevkTürü", comboBox3.Text);
                        komut.Parameters.AddWithValue("@SiparişiYazan", label59.Text);
                        komut.Parameters.AddWithValue("@M2KapakAdet", label30.Text);
                        komut.Parameters.AddWithValue("@OnayTarihi", Convert.ToDateTime(DateTime.Now.ToString()));
                        komut.Parameters.AddWithValue("@BID", textBox8.Text);
                        komut.Parameters.AddWithValue("@Nott", textBox11.Text);
                        komut.Parameters.AddWithValue("@SiparişTarihi", dateTimePicker1.Value);
                        komut.Parameters.AddWithValue("@TeslimTarihi", textBox10.Text);
                        komut.Parameters.AddWithValue("@DDS", label44.Text);
                        komut.ExecuteNonQuery();
                        MessageBox.Show("Bsr'ye satış eklenmiştir ve sipariş onaylanmıştır.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        liste();
                        bgl.baglanti().Close();
                        SqlConnection.ClearPool(bgl.baglanti());
                        textBox19.Text = DateTime.Now.ToString();
                        for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                        {
                            //m2
                            SqlCommand komut2 = new SqlCommand();
                            komut2.CommandText = "SELECT * FROM Stoklar where Malzeme=@Malzeme";
                            komut2.Parameters.AddWithValue("@Malzeme", dataGridView1.Rows[i].Cells["Renk"].Value.ToString());
                            komut2.Connection = bgl.baglanti();
                            komut2.CommandType = CommandType.Text;

                            SqlDataReader dr;
                            dr = komut2.ExecuteReader();
                            while (dr.Read())
                            {
                                mevcutstok = Convert.ToDouble(dr["MevcutStok"].ToString());
                                kullanılanstok = Convert.ToDouble(dr["Kullanılan"].ToString());
                            }

                            harcananstok = Convert.ToDouble(dataGridView1.Rows[i].Cells["M2"].Value.ToString()) * 2;
                            string sorgu2 = "UPDATE Stoklar SET Kullanılan=@p2, Kalan=@p3 WHERE Malzeme=@Malzeme";
                            SqlCommand cmd2;
                            cmd2 = new SqlCommand(sorgu2, bgl.baglanti());
                            cmd2.Parameters.AddWithValue("@Malzeme", dataGridView1.Rows[i].Cells["Renk"].Value.ToString());
                            cmd2.Parameters.AddWithValue("@p2", (kullanılanstok + harcananstok).ToString());
                            cmd2.Parameters.AddWithValue("@p3", (mevcutstok - kullanılanstok - harcananstok).ToString());
                            cmd2.ExecuteNonQuery();
                            bgl.baglanti().Close();
                            SqlConnection.ClearPool(bgl.baglanti());

                            //----------------------------------------------------------------------------------------------//mdf
                            SqlCommand komut3 = new SqlCommand();
                            komut3.CommandText = "SELECT * FROM Stoklar where Malzeme=@Malzeme";
                            komut3.Parameters.AddWithValue("@Malzeme", "18 MM TEKYÜZ MDF");
                            komut3.Connection = bgl.baglanti();
                            komut3.CommandType = CommandType.Text;

                            SqlDataReader dr3;
                            dr3 = komut3.ExecuteReader();
                            while (dr3.Read())
                            {
                                mevcutstokmdf = Convert.ToDouble(dr3["MevcutStok"].ToString());
                                kullanılanstokmdf = Convert.ToDouble(dr3["Kullanılan"].ToString());
                            }

                            harcananstokmdf = Convert.ToDouble(dataGridView1.Rows[i].Cells["M2"].Value.ToString()) / 5;
                            string sorgu3 = "UPDATE Stoklar SET Kullanılan=@Kullanılan, Kalan=@Kalan WHERE Malzeme=@Malzeme";
                            SqlCommand cmd3;
                            cmd3 = new SqlCommand(sorgu3, bgl.baglanti());
                            cmd3.Parameters.AddWithValue("@Malzeme", "18 MM TEKYÜZ MDF");
                            cmd3.Parameters.AddWithValue("@Kullanılan", (kullanılanstokmdf + harcananstokmdf).ToString());
                            cmd3.Parameters.AddWithValue("@Kalan", (mevcutstokmdf - kullanılanstokmdf - harcananstokmdf).ToString());
                            cmd3.ExecuteNonQuery();
                            bgl.baglanti().Close();
                            SqlConnection.ClearPool(bgl.baglanti());

                            //----------------------------------------------------------------------------------------------//tutkal
                            SqlCommand komut4 = new SqlCommand();
                            komut4.CommandText = "SELECT * FROM Stoklar where Malzeme=@Malzeme";
                            komut4.Parameters.AddWithValue("@Malzeme", "TUTKAL");
                            komut4.Connection = bgl.baglanti();
                            komut4.CommandType = CommandType.Text;

                            SqlDataReader dr4;
                            dr4 = komut4.ExecuteReader();
                            while (dr4.Read())
                            {
                                mevcutstoktutkal = Convert.ToDouble(dr4["MevcutStok"].ToString());
                                kullanılanstoktutkal = Convert.ToDouble(dr4["Kullanılan"].ToString());
                            }

                            harcananstoktutkal = Convert.ToDouble(dataGridView1.Rows[i].Cells["M2"].Value.ToString()) / 4;
                            string sorgu4 = "UPDATE Stoklar SET Kullanılan=@Kullanılan, Kalan=@Kalan WHERE Malzeme=@Malzeme";
                            SqlCommand cmd4;
                            cmd4 = new SqlCommand(sorgu4, bgl.baglanti());
                            cmd4.Parameters.AddWithValue("@Malzeme", "TUTKAL");
                            cmd4.Parameters.AddWithValue("@Kullanılan", (kullanılanstoktutkal + harcananstoktutkal).ToString());
                            cmd4.Parameters.AddWithValue("@Kalan", (mevcutstoktutkal - kullanılanstoktutkal - harcananstoktutkal).ToString());
                            cmd4.ExecuteNonQuery();
                            bgl.baglanti().Close();
                            SqlConnection.ClearPool(bgl.baglanti());
                            StokGetir();
                            TutkalStok();
                            MDFStok();
                        }

                        checkBox1.Visible = false;
                        button9.Visible = false;
                        button10.Visible = false;
                        button8.Visible = false;
                        button1.Enabled = false;
                        button2.Enabled = false;
                        button3.Enabled = false;
                        if (button3.Enabled == false)
                        {
                            button3.BackColor = Color.Green;
                        }
                        button3.Text = "ONAYLANDI";

                    }
                    bgl.baglanti().Close();
                    SqlConnection.ClearPool(bgl.baglanti());
                    Form3 frm = new Form3();
                    frm.yetki = yetki;
                    frm.kullaniciadi = kullaniciadi;
                    frm.hangiformdan = hangiformdan;
                    this.Hide();
                    frm.Show();
                }
            }
        }

        private void textBox17_TextChanged(object sender, EventArgs e)
        {

        }

        private void SiparişiYazan()
        {
            try
            {
                SqlCommand komut4 = new SqlCommand();
                komut4.CommandText = "SELECT * FROM Siparişler where SiparisNo=@SiparisNo";
                komut4.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
                komut4.Connection = bgl.baglanti();
                komut4.CommandType = CommandType.Text;

                SqlDataReader dr4;
                dr4 = komut4.ExecuteReader();
                while (dr4.Read())
                {
                    label60.Text = dr4["EkleyenKullanici"].ToString();
                }
            }
            catch (Exception)
            {

                throw;
            }

        }
        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                SiparişiYazan();
                Excel.Application excel = new Excel.Application();
                Excel.Workbook workbook = null;

                excel.Visible = false;
                object Missing = Type.Missing;
                //Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Modeks_Dosyalar\\SIPARIS.xlsx ");
                //Workbook workbook = excel.Workbooks.Open("C:\\Users\\Enes\\Desktop\\SIPARIS.xlsx ");


                if (label42.Text != "0")
                {
                    workbook = excel.Workbooks.Open(@"C:\Modeks_Dosyalar\SIPARIS.xlsx");
                }
                else
                {
                    workbook = excel.Workbooks.Open(@"C:\Modeks_Dosyalar\SIPARIS-Iskontosuz.xlsx");
                }

                Worksheet sheet1 = (Worksheet)workbook.Sheets[1];
                for (int k = 0; k < dataGridView1.Rows.Count - 1; k++)
                {
                    Range line = (Range)sheet1.Rows[11 + k];
                    line.Insert();
                }
                sheet1.Cells[3, 4].Value = textBox1.Text; // siparisno yazdırma
                sheet1.Cells[4, 4].Value = comboBox1.Text; // müşteri yazdırma
                sheet1.Cells[5, 4].Value = textBox3.Text; // adres yazdırma
                sheet1.Cells[7, 4].Value = textBox4.Text; // telefon yazdırma
                sheet1.Cells[7, 6].Value = textBox4.Text; // telefon yazdırma

                sheet1.Cells[7, 9].Value = dateTimePicker1.Text; // sip tarih yazdırma
                sheet1.Cells[8, 9].Value = textBox19.Text; // onay tarih yazdırma
                sheet1.Cells[9, 9].Value = textBox10.Text; // tes tarih yazdırma

                sheet1.Cells[12 + dataGridView1.Rows.Count - 1, 9].Value = label28.Text; // adet toplam yazdırma
                sheet1.Cells[12 + dataGridView1.Rows.Count - 1, 10].Value = label27.Text; // toplam m2 yazdırma
                sheet1.Cells[16 + dataGridView1.Rows.Count - 1, 5].Value = label74.Text; // Ara Toplam yazdırma
                if (label42.Text != "0")
                {
                    sheet1.Cells[17 + dataGridView1.Rows.Count - 1, 5].Value = label42.Text;
                    sheet1.Cells[18 + dataGridView1.Rows.Count - 1, 5].Value = label76.Text; // Kargo yazdırma
                    sheet1.Cells[19 + dataGridView1.Rows.Count - 1, 5].Value = label44.Text; // DDS yazdırma
                    sheet1.Cells[20 + dataGridView1.Rows.Count - 1, 5].Value = label63.Text; // toplam tutar yazdırma
                } else
                {
                    sheet1.Cells[17 + dataGridView1.Rows.Count - 1, 5].Value = label76.Text; // Kargo yazdırma
                    sheet1.Cells[18 + dataGridView1.Rows.Count - 1, 5].Value = label44.Text; // DDS yazdırma
                    sheet1.Cells[19 + dataGridView1.Rows.Count - 1, 5].Value = label63.Text; // toplam tutar yazdırma
                }
               
                sheet1.Cells[8, 4].Value = label60.Text; // toplam tutar yazdırma
                sheet1.Cells[3, 6].Value = textBox8.Text; // BID yazdırma

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        if (j != 0 && j != 3 && j != 9 && j != 10 && j != 11 && j != 12)
                        {
                            sheet1.Cells[i + 11, j + 2].Value = dataGridView1.Rows[i].Cells[j].Value;
                        }
                    }
                }
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string pdfFilePath = Path.Combine(desktopPath, "" + textBox1.Text + " - SIP " + comboBox1.Text + ".pdf");

                // Belgeyi PDF olarak kaydet
                sheet1.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, pdfFilePath);
                MessageBox.Show("Pdf dosyası masaüstüne oluşturulmuştur.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //sheet1.PrintPreview();
                workbook.Close(false);
                excel.Quit();
                foreach (var process in Process.GetProcessesByName("EXCEL"))
                {
                    process.Kill();
                }
                Marshal.ReleaseComObject(sheet1);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excel);

                // PDF dosyasını açabilirsiniz
                System.Diagnostics.Process.Start(pdfFilePath);

                // Excel'i görev yöneticisinden kapat
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet1);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Müşteri Sipariş Formu Hatası :" + ex.Message);
            }

        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                SiparişiYazan();
                Excel.Application excel = new Excel.Application();
                excel.Visible = false;
                object Missing = Type.Missing;
                //Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Modeks_Dosyalar\\ÜretimFormu.xlsx ");
                //Workbook workbook = excel.Workbooks.Open("C:\\Users\\Enes\\Desktop\\ÜretimFormu.xlsx ");
                Excel.Workbook workbook = excel.Workbooks.Open("C:\\Modeks_Dosyalar\\ÜretimFormu.xlsx");
                Worksheet sheet2 = (Worksheet)workbook.Sheets[1];
                for (int k = 0; k < dataGridView1.Rows.Count - 1; k++)
                {
                    Range line = (Range)sheet2.Rows[11 + k];
                    line.Insert();
                }
                sheet2.Cells[3, 4].Value = textBox1.Text; // siparisno yazdırma
                sheet2.Cells[4, 4].Value = comboBox1.Text; // müşteri yazdırma
                sheet2.Cells[5, 4].Value = textBox3.Text; // adres yazdırma
                sheet2.Cells[7, 4].Value = textBox4.Text; // telefon yazdırma
                sheet2.Cells[7, 6].Value = textBox4.Text; // telefon yazdırma
                sheet2.Cells[7, 9].Value = dateTimePicker1.Text; // sip tarih yazdırma
                sheet2.Cells[8, 9].Value = textBox19.Text; // onay tarih yazdırma
                sheet2.Cells[9, 9].Value = textBox10.Text; // tes tarih yazdırma
                sheet2.Cells[12 + dataGridView1.Rows.Count - 1, 9].Value = label28.Text; // adet toplam yazdırma
                sheet2.Cells[12 + dataGridView1.Rows.Count - 1, 10].Value = label27.Text; // toplam m2 yazdırma
                sheet2.Cells[8, 4].Value = label60.Text; // toplam tutar yazdırma
                sheet2.Cells[5, 7].Value = comboBox2.Text;

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        if (j != 0 && j != 3 && j != 9 && j != 10 && j != 11 && j != 12)
                            sheet2.Cells[i + 11, j + 2].Value = dataGridView1.Rows[i].Cells[j].Value;
                    }
                }


                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string pdfFilePath = Path.Combine(desktopPath, "" + textBox1.Text + " - URT " + comboBox1.Text + ".pdf");

                // Belgeyi PDF olarak kaydet
                sheet2.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, pdfFilePath);
                MessageBox.Show("Pdf dosyası masaüstüne oluşturulmuştur", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);

                workbook.Close(false);
                excel.Quit();
                foreach (var process in Process.GetProcessesByName("EXCEL"))
                {
                    process.Kill();
                }
                // Excel'i görev yöneticisinden kapat
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet2);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception)
            {

                throw;
            }

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != "" && comboBox6.Text != "")
            {
                Form15 frm = new Form15();
                frm.müşteri = comboBox1.Text;
                frm.renk = comboBox6.Text;
                frm.Show();
            }
            else
            {
                MessageBox.Show("Lütfen müşteri ve renk seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {
            //if (textBox15.Text == "0")
            //{
            //    MessageBox.Show("Fiyatı 0 girdiniz. Lütfen bir açıklama giriniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }
        private void KaydetVeÇık()
        {
            try
            {
                if (comboBox1.Text == "" || comboBox2.Text == "" || comboBox3.Text == "")
                {
                    MessageBox.Show("Müşteri Kısmını, Sipariş Şeklini ve Sipariş Türünü Seçiniz!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    KapakAdetToplat();
                    TasarımUcretıToplat();
                    KapakToplam();
                    M2KapakSayısı();
                    if (comboBox2.Text == "Acil")
                    {
                        AcilHesaplama();
                    }
                    M2KapakFarkıHesaplama();
                    KargoHesaplama();
                    Toplam();
                    AraToplam();
                    DDSHesaplama();
                    GenelToplam();

                    string teslimtarihi = Convert.ToDateTime(textBox10.Text).ToString("dd.MM.yyyy HH:mm:ss");
                    string sorgu2 = "UPDATE Siparişler SET ToplamFiyat=@ToplamFiyat,Müşteri=@Müşteri,SiparişTarihi=@SiparişTarihi,TeslimTarihi=@TeslimTarihi,İskonto=@İskonto WHERE SiparisNo=@SiparisNo";
                    SqlCommand cmd2;
                    cmd2 = new SqlCommand(sorgu2, bgl.baglanti());
                    cmd2.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
                    cmd2.Parameters.AddWithValue("@ToplamFiyat", label63.Text);
                    cmd2.Parameters.AddWithValue("@Müşteri", comboBox1.Text);
                    cmd2.Parameters.AddWithValue("@SiparişTarihi", dateTimePicker1.Value);
                    cmd2.Parameters.AddWithValue("@TeslimTarihi", teslimtarihi);
                    if (textBox16.Text == "" || textBox16.Text == null)
                        textBox16.Text = "0";
                    cmd2.Parameters.AddWithValue("@İskonto", textBox16.Text);
                    cmd2.ExecuteNonQuery();
                    bgl.baglanti().Close();
                    SqlConnection.ClearPool(bgl.baglanti());

                    MessageBox.Show("Sipariş kaydedilmiştir.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                        Form3 yeni = new Form3();
                        yeni.yetki = yetki;
                        yeni.kullaniciadi = kullaniciadi;
                        this.Hide();
                        yeni.Show();
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }

        }
        private void button4_Click_2(object sender, EventArgs e)
        {
            KaydetVeÇık();
        }

        private void Form10_FormClosed(object sender, FormClosedEventArgs e)
        {

            System.Windows.Forms.Application.Exit();
        }

        private void label66_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.Text == "Kargo")
            {
                //textBox10.Text = DateTime.Now.AddDays(20).ToString("dd.MM.yyyy HH:dd:ss");
                //liste();
                M2Toplat();
                KapakAdetToplat();
                TasarımUcretıToplat();
                KapakToplam();
                M2KapakSayısı();
                AcilHesaplama();
                M2KapakFarkıHesaplama();
                KargoHesaplama();
                Toplam();
                AraToplam();
                DDSHesaplama();
                GenelToplam();

                string sorgu = "UPDATE Siparişler SET Kargo=@Kargo WHERE SiparisNo=@SiparisNo";
                SqlCommand komut;
                komut = new SqlCommand(sorgu, bgl.baglanti());
                komut.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
                komut.Parameters.AddWithValue("@Kargo", label40.Text);
                komut.ExecuteNonQuery();

                string sorgu2 = "UPDATE Siparişler SET SevkTürü=@SevkTürü WHERE SiparisNo=@SiparisNo";
                SqlCommand komut2;
                komut2 = new SqlCommand(sorgu2, bgl.baglanti());
                komut2.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
                komut2.Parameters.AddWithValue("@SevkTürü", comboBox3.Text);
                komut2.ExecuteNonQuery();
            }
            else if (comboBox3.Text == "Fabrika")
            {
                //liste();
                M2Toplat();
                KapakAdetToplat();
                TasarımUcretıToplat();
                KapakToplam();
                M2KapakSayısı();
                AcilHesaplama();
                M2KapakFarkıHesaplama();
                KargoHesaplama();
                Toplam();
                AraToplam();
                DDSHesaplama();
                GenelToplam();


                string sorgu = "UPDATE Siparişler SET Kargo=@Kargo WHERE SiparisNo=@SiparisNo";
                SqlCommand komut;
                komut = new SqlCommand(sorgu, bgl.baglanti());
                komut.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
                komut.Parameters.AddWithValue("@Kargo", label40.Text);
                komut.ExecuteNonQuery();

                string sorgu2 = "UPDATE Siparişler SET SevkTürü=@SevkTürü WHERE SiparisNo=@SiparisNo";
                SqlCommand komut2;
                komut2 = new SqlCommand(sorgu2, bgl.baglanti());
                komut2.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
                komut2.Parameters.AddWithValue("@SevkTürü", comboBox3.Text);
                komut2.ExecuteNonQuery();
            }
        }

        private void label53_Click(object sender, EventArgs e)
        {

        }

        private void label52_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text == "Acil")
            {
                textBox10.Text = dateTimePicker1.Value.AddDays(5).ToString("dd.MM.yyyy HH:dd:ss");
            }
            else if (comboBox2.Text == "Normal")
            {
                textBox10.Text = dateTimePicker1.Value.AddDays(20).ToString("dd.MM.yyyy HH:dd:ss");
            }
        }

        private void label63_TextChanged(object sender, EventArgs e)
        {
            //iskontogetir();
            İskontoGüncelle();
        }

        private void label63_Click(object sender, EventArgs e)
        {

        }

        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void comboBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void button12_Click(object sender, EventArgs e)
        {
            Form22 frm = new Form22();
            frm.ShowDialog();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                if (hangiformdanModeksEski == "true")
                {
                    Form24 frm = new Form24();
                    this.Hide();
                    frm.ShowDialog();
                }
                else
                {
                    if (dataGridView1.Rows.Count > 1)
                    {
                        if (comboBox1.Text == "" || comboBox2.Text == "" || comboBox3.Text == "")
                        {
                            MessageBox.Show("Müşteri Kısmını, Sipariş Şeklini ve Sipariş Türünü Seçiniz!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else
                        {
                            KaydetVeÇık();
                        }
                    }
                    else
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
                            Form3 yeni = new Form3();
                            yeni.yetki = yetki;
                            yeni.kullaniciadi = kullaniciadi;
                            this.Hide();
                            yeni.Show();
                        }
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != "" && comboBox6.Text != "")
            {
                Form15 frm = new Form15();
                frm.müşteri = comboBox1.Text;
                frm.renk = comboBox6.Text;
                frm.Show();
            }
            else
            {
                MessageBox.Show("Lütfen müşteri ve renk seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        void birlesikliste()
        {
            string kayit = @"
    SELECT 
        CAST(id AS VARCHAR) as 'ID', 
        Model, 
        Renk, 
        TasarımÜcreti, 
        Özellik, 
        TRY_CAST(REPLACE(Boy, ',', '.') AS decimal(12, 1)) AS Boy,
	    TRY_CAST(REPLACE(En, ',', '.') AS decimal(12, 1)) AS En,
        Adet, 
        M2, 
        M2Fiyat, 
        Fiyat2, 
        BID, 
        Nott as 'Not',
        ISNULL(EkleyenKullanici,'') as EkleyenKullanici
    FROM 
        Siparişler 
    WHERE 
        SiparisNo = @SiparisNo ORDER BY 
	 CASE WHEN Özellik = 'CEKMECE BUTUN' THEN TRY_CAST(REPLACE(Boy, ',', '.') as decimal(12,4))
	 ELSE NULL END ,
	 CASE WHEN Özellik <> 'CEKMECE BUTUN' THEN 1
	 ELSE 0
	 END,
	 Boy DESC, 
	 En DESC";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            komut.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
            SqlDataAdapter da = new SqlDataAdapter(komut);
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);

            System.Data.DataTable groupedDt = new System.Data.DataTable();
            groupedDt.Columns.Add("ID", typeof(string));
            groupedDt.Columns.Add("Model", typeof(string));
            groupedDt.Columns.Add("Renk", typeof(string));
            groupedDt.Columns.Add("TasarımÜcreti", typeof(decimal));
            groupedDt.Columns.Add("Özellik", typeof(string));
            groupedDt.Columns.Add("Boy", typeof(decimal));
            groupedDt.Columns.Add("En", typeof(decimal));
            groupedDt.Columns.Add("Adet", typeof(int));
            groupedDt.Columns.Add("M2", typeof(decimal));
            groupedDt.Columns.Add("M2Fiyat", typeof(decimal));
            groupedDt.Columns.Add("Fiyat2", typeof(decimal));
            groupedDt.Columns.Add("BID", typeof(int));
            groupedDt.Columns.Add("Not", typeof(string));
            groupedDt.Columns.Add("EkleyenKullanici", typeof(string));

            var groupedData = dt.AsEnumerable()
    .GroupBy(row => new
    {
        Model = row.Field<string>("Model"),
        Renk = row.Field<string>("Renk"),
        Özellik = row.Field<string>("Özellik"),
        Boy = decimal.TryParse(row["Boy"].ToString(), out var boy) ? boy : (decimal?)null,
        En = decimal.TryParse(row["En"].ToString(), out var en) ? en : (decimal?)null
    })
    .Select(g => new
    {
        ID = g.First().Field<string>("ID"),
        Model = g.Key.Model,
        Renk = g.Key.Renk,
        Özellik = g.Key.Özellik,
        TasarımÜcreti = g.Sum(row => decimal.TryParse(row["TasarımÜcreti"].ToString(), out var tasarimUcreti) ? tasarimUcreti : 0),
        Boy = g.Key.Boy ?? 0,
        En = g.Key.En ?? 0,
        Adet = g.Sum(row => int.TryParse(row["Adet"].ToString(), out var adet) ? adet : 0),
        M2 = g.Sum(row => decimal.TryParse(row["M2"].ToString(), out var m2) ? m2 : 0),
        M2Fiyat = decimal.TryParse(g.First().Field<string>("M2Fiyat"), out var m2Fiyat) ? m2Fiyat : 0, 
        Fiyat2 = g.Sum(row => decimal.TryParse(row["Fiyat2"].ToString(), out var fiyat2) ? fiyat2 : 0),
        BID = g.Sum(row => int.TryParse(row["BID"].ToString(), out var bid) ? bid : 0),
        Not = g.First().Field<string>("Not"),
        EkleyenKullanici = g.First().Field<string>("EkleyenKullanici")

    })
    .ToList();


            foreach (var item in groupedData)
            {
                groupedDt.Rows.Add(item.ID, item.Model, item.Renk, item.TasarımÜcreti, item.Özellik, item.Boy, item.En, item.Adet, item.M2, item.M2Fiyat, item.Fiyat2, item.BID, item.Not, item.EkleyenKullanici);
            }

            dataGridView4.DataSource = groupedDt;
        }
        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if(checkBox3.Checked)
            {
                birlesikliste();
                dataGridView4.Visible = true;
                dataGridView1.Visible = false;
                checkBox1.Checked = false;
            } else
            {
                liste();
                dataGridView4.Visible = false;
                dataGridView1.Visible = true;
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
         
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
           
        }

        private void button14_Click_1(object sender, EventArgs e)
        {
            try
            {
                Type formType = Type.GetType($"Modeks.{hangiformdan}");

                if (formType != null)
                {
                    Form frm = (Form)Activator.CreateInstance(formType);

                    dynamic dynForm = frm;

                    dynForm.yetki = yetki;
                    this.Hide();
                    frm.ShowDialog();
                    this.Show();
                }
                else
                {
                    MessageBox.Show("Form bulunamadı: " + hangiformdan);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Bir hata oluştu: " + ex.Message);
            }
        }


        private void label77_Click(object sender, EventArgs e)
        {


        }
    }
}

//ctrl+m toplu daraltma