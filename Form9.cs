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
using System.IO;

namespace Modeks
{
    public partial class Form9 : Form
    {
        public Form9()
        {
            InitializeComponent();
        }
        sqlsinif bgl = new sqlsinif();
        public string yetki;
        double yapılacakkargom2 = 0;
        double yapılacakkargoadet = 0;
        double yapılacakfabm2 = 0;
        double yapılacakfabadet = 0;
        double toplamm2 = 0;
        double toplamadet = 0;
        double yapılanm2 = 0;
        double yapılanadet = 0;
        double yapılanfabkargom2 = 0;
        double yapılanfabkargoadet = 0;

        double bugünbasılankargom2 = 0;
        double bugünbasılankargoadet = 0;

        double bugünbasılanfabm2 = 0;
        double bugünbasılanfabadet = 0;

        double bugünbasılantoplamm2 = 0;
        double bugünbasılantoplamadet = 0;
        private void liste()
        {
            string kayit = "SELECT DISTINCT SiparisNo,Müşteri,Renk,SiparişTipi,ToplamM2,ToplamAdet,SevkTürü,Model,SiparişTarihi,TeslimTarihi,BID,Adres,PaketTarihi From Siparişler where AnaSiparişMi=@p1 AND Aşama=@Aşama ORDER BY SiparisNo DESC";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            komut.Parameters.AddWithValue("@p1", "Evet");
            komut.Parameters.AddWithValue("@Aşama", "Kargo");
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            onayrenklendir();
        }
        private void liste2()
        {
            string kayit = "SELECT DISTINCT SiparisNo,Müşteri,Renk,SiparişTipi,ToplamM2,ToplamAdet,SevkTürü,Model,SiparişTarihi,TeslimTarihi,BID,Adres,PaketTarihi,Telefon From Siparişler where AnaSiparişMi=@p1 AND Aşama=@Aşama ORDER BY SiparisNo ASC";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            komut.Parameters.AddWithValue("@p1", "Evet");
            komut.Parameters.AddWithValue("@Aşama", "Kargo");
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView2.DataSource = dt;
            onayrenklendir();
        }
        private void YapılacakKargoM2()
        {
            yapılacakkargom2 = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true && dataGridView1.Rows[i].Cells[6].Value.ToString() == "Kargo")
                    yapılacakkargom2 += Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value);
            }
            textBox5.Text = yapılacakkargom2.ToString("0.##");
        }
        private void YapılacakKargoAdet()
        {
            yapılacakkargoadet = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true && dataGridView1.Rows[i].Cells[6].Value.ToString() == "Kargo")
                    yapılacakkargoadet += Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value);
            }
            textBox6.Text = yapılacakkargoadet.ToString("0.##");
        }
        private void YapılacakFabM2()
        {
            yapılacakfabm2 = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true && dataGridView1.Rows[i].Cells[6].Value.ToString() == "Fabrika")
                    yapılacakfabm2 += Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value);
            }
            textBox8.Text = yapılacakfabm2.ToString("0.##");
        }
        private void YapılacakFabAdet()
        {
            yapılacakfabadet = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                if (dataGridView1.Rows[i].Visible == true && dataGridView1.Rows[i].Cells[6].Value.ToString() == "Fabrika")
                    yapılacakfabadet += Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value);
            }
            textBox7.Text = yapılacakfabadet.ToString("0.##");
        }
        private void ToplamM2Adet()
        {
            toplamm2 = 0;
            toplamadet = 0;
            toplamm2 = yapılacakkargom2 + yapılacakfabm2;
            toplamadet = yapılacakkargoadet + yapılacakfabadet;
            textBox10.Text = toplamm2.ToString("0.##");
            textBox9.Text = toplamadet.ToString("0.##");
        }
        private void YapılanPaketKargo()
        {
            yapılanm2 = 0;
            yapılanadet = 0;
            yapılanfabkargom2 = 0;
            yapılanfabkargoadet = 0;
            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT *FROM Siparişler where Aşama=@p1 AND SevkTürü=@p2";
            komut.Parameters.AddWithValue("@p1", "Hazır");
            komut.Parameters.AddWithValue("@p2", "Kargo");
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                yapılanm2 += Convert.ToDouble(dr["ToplamM2"]);
                yapılanadet += Convert.ToDouble(dr["ToplamAdet"]);

                textBox3.Text = yapılanm2.ToString();
                textBox4.Text = yapılanadet.ToString();
                textBox14.Text = yapılacakkargom2.ToString("0.##");
                textBox13.Text = yapılacakkargoadet.ToString("0.##");
            }
        }
        private void YapılanFabKargo()
        {
            yapılanfabkargom2 = 0;
            yapılanfabkargoadet = 0;
            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT *FROM Siparişler where Aşama=@p1";
            komut.Parameters.AddWithValue("@p1", "Hazır");
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                yapılanfabkargom2 += Convert.ToDouble(dr["ToplamM2"]);
                yapılanfabkargoadet += Convert.ToDouble(dr["ToplamAdet"]);

                textBox14.Text = yapılanfabkargom2.ToString("0.##");
                textBox13.Text = yapılanfabkargoadet.ToString("0.##");
            }
        }
        private void Methodlar()
        {
            liste2();
            siparisnogetir();
            müşterigetir();
            aynısipnugetirme();
            YapılacakKargoM2();
            YapılacakKargoAdet();
            YapılacakFabM2();
            YapılacakFabAdet();
            ToplamM2Adet();
            YapılanPaketKargo();
            YapılanFabKargo();
            BugünBasılanlar();
            AcilSipariş();
            TeslimTarihine3GünKalanlarıYakSöndür();
            onayrenklendir();
        }
        string sevktürü;
        DateTime tarih;
        private void BugünBasılanlar()
        {
            bugünbasılankargom2 = 0;
            bugünbasılankargoadet = 0;
            bugünbasılanfabm2 = 0;
            bugünbasılanfabadet = 0;
            bugünbasılantoplamm2 = 0;
            bugünbasılantoplamadet = 0;

            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT *FROM Siparişler where AnaSiparişMi=@AnaSiparişMi";
            komut.Parameters.AddWithValue("@AnaSiparişMi", "Evet");
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                sevktürü = dr["SevkTürü"].ToString();
                if (dr["PaketTarihi"].ToString() != "")
                {
                    tarih = Convert.ToDateTime(dr["PaketTarihi"].ToString());
                    if (sevktürü == "Kargo" && tarih.ToString("yyyy - MM - dd") == (DateTime.Now.ToString("yyyy - MM - dd")))
                    {
                        bugünbasılankargom2 += Convert.ToDouble(dr["M2"]);
                        bugünbasılankargoadet += Convert.ToDouble(dr["Adet"]);

                        textBox2.Text = bugünbasılankargom2.ToString();
                        textBox26.Text = bugünbasılankargom2.ToString();
                        textBox25.Text = bugünbasılankargoadet.ToString();

                    }
                    else if (sevktürü == "Fabrika" && tarih.ToString("yyyy - MM - dd") == (DateTime.Now.ToString("yyyy - MM - dd")))
                    {
                        bugünbasılanfabm2 += Convert.ToDouble(dr["M2"]);
                        bugünbasılanfabadet += Convert.ToDouble(dr["Adet"]);
                        textBox1.Text = bugünbasılanfabm2.ToString();
                        textBox24.Text = bugünbasılanfabm2.ToString();
                        textBox23.Text = bugünbasılanfabadet.ToString();
                    }
                }
            }
            bugünbasılantoplamm2 += bugünbasılankargom2 + bugünbasılanfabm2;
            bugünbasılantoplamadet += bugünbasılankargoadet + bugünbasılanfabadet;
            textBox33.Text = bugünbasılantoplamm2.ToString();
            textBox22.Text = bugünbasılantoplamm2.ToString();
            textBox34.Text = bugünbasılantoplamadet.ToString();
            textBox21.Text = bugünbasılantoplamadet.ToString();
            textBox35.Text = (bugünbasılantoplamadet / bugünbasılantoplamm2).ToString();
        }

        private void siparisnogetir()
        {
            comboBox3.Items.Clear();
            comboBox1.Items.Clear();
            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT DISTINCT SiparisNo FROM Siparişler where AnaSiparişMi=@p1 AND Aşama=@Aşama";
            komut.Parameters.AddWithValue("@p1", "Evet");
            komut.Parameters.AddWithValue("@Aşama", "Kargo");
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                comboBox3.Items.Add(dr["SiparisNo"]);
                comboBox1.Items.Add(dr["SiparisNo"]);
            }
        }
        private void müşterigetir()
        {
            comboBox2.Items.Clear();
            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT DISTINCT Müşteri FROM Siparişler where AnaSiparişMi=@p1 AND Aşama=@Aşama";
            komut.Parameters.AddWithValue("@p1", "Evet");
            komut.Parameters.AddWithValue("@Aşama", "Kargo");
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                comboBox2.Items.Add(dr["Müşteri"]);
            }
        }
        private void siparisnoyagöresırala()
        {
            string srg = comboBox1.Text;
            string sorgu = "SELECT DISTINCT SiparisNo,Müşteri,Renk,SiparişTipi,ToplamM2,ToplamAdet,SevkTürü,Model,SiparişTarihi,TeslimTarihi,BID,Adres From Siparişler where SiparisNo Like '" + srg + "' AND AnaSiparişMi='" + "Evet" + "' AND Aşama='" + "Kargo" + "' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            onayrenklendir();
        }
        private void müşteriyegöresırala()
        {
            string srg = comboBox2.Text;
            string sorgu = "SELECT DISTINCT SiparisNo,Müşteri,Renk,SiparişTipi,ToplamM2,ToplamAdet,SevkTürü,Model,SiparişTarihi,TeslimTarihi,BID,Adres From Siparişler where Müşteri Like '" + srg + "' AND AnaSiparişMi='" + "Evet" + "' AND Aşama='" + "Kargo" + "' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            onayrenklendir();
        }

        string acilsipariş;
        string kesilditarihi;
        string Etiket;
        string membrantarihi;
        string pakettarihi;
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
                pakettarihi = dr["PaketTarihi"].ToString();
                if ((acilsipariş == "Acil" && kesilditarihi.Length < 2) || (acilsipariş == "Acil" && Etiket.Length < 2) || (acilsipariş == "Acil" && membrantarihi.Length < 2) || (acilsipariş == "Acil" && pakettarihi.Length < 2))
                {
                    timer2.Start();
                    label9.Text = "Dikkat! Acil Sipariş Var! Dikkat! Acil Sipariş Var! Dikkat! Acil Sipariş Var!";
                    label9.BackColor = Color.Red;
                    break;
                }
                else
                {
                    timer2.Stop();
                    label9.Text = "YAPILACAK PAKETLER";
                    label9.TextAlign = ContentAlignment.MiddleCenter;
                    label9.BackColor = Color.Blue;
                }
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
        private void onayrenklendir()
        {
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
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
                if (Convert.ToString(dataGridView1.Rows[i].Cells["SevkTürü"].Value.ToString()) == "Kargo")
                {
                    dataGridView1.Rows[i].Cells["SevkTürü"].Style.BackColor = Color.CadetBlue;
                    dataGridView1.Rows[i].Cells["SevkTürü"].Style.ForeColor = Color.White;
                }
                if (Convert.ToString(dataGridView1.Rows[i].Cells["SevkTürü"].Value.ToString()) == "Fabrika")
                {
                    dataGridView1.Rows[i].Cells["SevkTürü"].Style.BackColor = Color.Orange;
                    dataGridView1.Rows[i].Cells["SevkTürü"].Style.ForeColor = Color.White;
                }
            }
        }
        private void Form9_Load(object sender, EventArgs e)
        {
            if (yetki == "Paket ve Kargo")
            {
                button28.Visible = false;
            }
            timer1.Start();
            liste();
            aynısipnugetirme();
            dataGridView1.Columns["Model"].Visible = false;
            dataGridView1.Columns["SiparişTarihi"].Visible = false;
            dataGridView1.Columns["TeslimTarihi"].Visible = false;
            dataGridView1.Columns["BID"].Visible = false;
            dataGridView1.Columns["Adres"].Visible = false;
            Methodlar();
            radioButton1.Checked = true;
        }

        private void pictureBox8_Click(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            label2.Text = DateTime.Now.ToLongDateString();
            label12.Text = DateTime.Now.ToLongTimeString();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        string izin;
        private void button5_Click(object sender, EventArgs e)
        {
            if (comboBox3.Text != "" && textBox11.Text != "")
            {
                for (int y = 0; y < dataGridView1.Rows.Count - 1; y++)
                {
                    if (dataGridView1.Rows[y].Cells[0].Value.ToString() == comboBox3.Text)
                    {
                        if (dataGridView1.Rows[y].Cells["PaketTarihi"].Value.ToString().Length > 2)
                        {
                            DialogResult d1 = new DialogResult();
                            d1 = MessageBox.Show("Bu paket daha önce çıkarılmış. Tekrar çıkartmak istiyor musunuz ?", "Uyarı", MessageBoxButtons.YesNo);
                            if (d1 == DialogResult.Yes)
                            {
                                if (radioButton1.Checked == true)
                                {

                                    int j = 2;
                                    Excel.Application excel = new Excel.Application();
                                    excel.Visible = true;
                                    object Missing = Type.Missing;
                                    //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Modeks_Dosyalar\\KARGO.xlsx ");
                                    //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Enes\\Desktop\\KARGO.xlsx");
                                    Excel.Workbook workbook = excel.Workbooks.Open("C:\\Modeks_Dosyalar\\KARGO.xlsx");
                                    Excel.Worksheet sheet1 = (Excel.Worksheet)workbook.Sheets[1];
                                    sheet1.Cells[4, 1].Value = dataGridView1.Rows[0].Cells[0].Value.ToString();
                                    sheet1.Cells[6, 1].Value = dataGridView1.Rows[0].Cells[1].Value.ToString();
                                    sheet1.Cells[12, 1].Value = dataGridView1.Rows[0].Cells[7].Value.ToString();
                                    sheet1.Cells[13, 1].Value = dataGridView1.Rows[0].Cells[2].Value.ToString();
                                    sheet1.Cells[14, 1].Value = dataGridView1.Rows[0].Cells[8].Value.ToString();
                                    sheet1.Cells[15, 1].Value = dataGridView1.Rows[0].Cells[9].Value.ToString();
                                    sheet1.Cells[14, 4].Value = 1 + "/" + textBox11.Text;
                                    sheet1.Cells[8, 1].Value = dataGridView1.Rows[0].Cells[11].Value.ToString();
                                    sheet1.Cells[4, 3].Value = dataGridView1.Rows[0].Cells[10].Value.ToString();
                                    for (int i = 0; i < Convert.ToInt32(textBox11.Text) - 1; i++)
                                    {
                                        sheet1.Range["A1:D15"].Copy(sheet1.Range["A" + (j + 14) + ""]);
                                        j += 15;
                                        sheet1.Cells[(12 + j), 4].Value = (i + 2) + "/" + textBox11.Text;

                                    }
                                    //sheet1.Range["A14:F23"].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                                    //sheet1.PrintPreview();
                                    //workbook.Close(false);
                                    //excel.Quit();
                                    string etiketFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Etiket");

                                    if (!Directory.Exists(etiketFolderPath))
                                    {
                                        Directory.CreateDirectory(etiketFolderPath);
                                    }

                                    string excelFilePath = Path.Combine(etiketFolderPath, "Paket" + dataGridView1.Rows[0].Cells[0].Value.ToString() + ".xlsx");


                                    workbook.SaveAs(excelFilePath);

                                    sheet1.PrintOutEx(Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing);
                                    workbook.Close(false, Missing, Missing);
                                    excel.Quit();

                                    string tarih = Convert.ToDateTime(DateTime.Now).ToString("yyyy-MM-dd HH:mm:ss");

                                    string sorgu4 = "UPDATE Siparişler SET PaketTarihi=@PaketTarihi, PaketSayısı=@PaketSayısı, Aşama=@Aşama WHERE SiparisNo=@SiparisNo AND AnaSiparişMi='" + "Evet" + "' ";
                                    SqlCommand komut4;
                                    komut4 = new SqlCommand(sorgu4, bgl.baglanti());
                                    komut4.Parameters.AddWithValue("@SiparisNo", comboBox3.Text);
                                    komut4.Parameters.AddWithValue("@PaketTarihi", tarih);
                                    komut4.Parameters.AddWithValue("@PaketSayısı", textBox11.Text);
                                    komut4.Parameters.AddWithValue("@Aşama", "Hazır");
                                    komut4.ExecuteNonQuery();
                                    label19.Text = comboBox3.Text;
                                    label14.Text = dataGridView1.Rows[0].Cells["Müşteri"].Value.ToString();
                                    liste();
                                    Methodlar();
                                }
                                else if (radioButton2.Checked == true)
                                {
                                    int j = 2;
                                    Excel.Application excel = new Excel.Application();
                                    excel.Visible = true;
                                    object Missing = Type.Missing;
                                    //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Enes\\Desktop\\KARGO.xlsx");
                                    //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Modeks_Dosyalar\\KARGO.xlsx ");
                                    Excel.Workbook workbook = excel.Workbooks.Open("C:\\Modeks_Dosyalar\\KARGO.xlsx");
                                    Excel.Worksheet sheet1 = (Excel.Worksheet)workbook.Sheets[2];
                                    sheet1.Cells[4, 1].Value = dataGridView1.Rows[0].Cells[0].Value.ToString();
                                    sheet1.Cells[6, 1].Value = dataGridView1.Rows[0].Cells[1].Value.ToString();
                                    sheet1.Cells[12, 1].Value = dataGridView1.Rows[0].Cells[7].Value.ToString();
                                    sheet1.Cells[13, 1].Value = dataGridView1.Rows[0].Cells[2].Value.ToString();
                                    sheet1.Cells[14, 1].Value = dataGridView1.Rows[0].Cells[8].Value.ToString();
                                    sheet1.Cells[15, 1].Value = dataGridView1.Rows[0].Cells[9].Value.ToString();
                                    sheet1.Cells[14, 4].Value = 1 + "/" + textBox11.Text;
                                    sheet1.Cells[8, 1].Value = dataGridView1.Rows[0].Cells[11].Value.ToString();
                                    sheet1.Cells[4, 3].Value = dataGridView1.Rows[0].Cells[10].Value.ToString();
                                    for (int i = 0; i < Convert.ToInt32(textBox11.Text) - 1; i++)
                                    {
                                        sheet1.Range["A1:D15"].Copy(sheet1.Range["A" + (j + 14) + ""]);
                                        j += 15;
                                        sheet1.Cells[(12 + j), 4].Value = (i + 2) + "/" + textBox11.Text;

                                    }
                                    //sheet1.Range["A14:F23"].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                                    string etiketFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Etiket");

                                    if (!Directory.Exists(etiketFolderPath))
                                    {
                                        Directory.CreateDirectory(etiketFolderPath);
                                    }

                                    string excelFilePath = Path.Combine(etiketFolderPath, "Paket" + dataGridView1.Rows[0].Cells[0].Value.ToString() + ".xlsx");


                                    workbook.SaveAs(excelFilePath);
                                    sheet1.PrintOutEx(Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing);
                                    workbook.Close(false, Missing, Missing);
                                    excel.Quit();

                                    string tarih = Convert.ToDateTime(DateTime.Now).ToString("yyyy-MM-dd HH:mm:ss");

                                    string sorgu4 = "UPDATE Siparişler SET PaketTarihi=@PaketTarihi, PaketSayısı=@PaketSayısı, Aşama=@Aşama WHERE SiparisNo=@SiparisNo AND AnaSiparişMi='" + "Evet" + "' ";
                                    SqlCommand komut4;
                                    komut4 = new SqlCommand(sorgu4, bgl.baglanti());
                                    komut4.Parameters.AddWithValue("@SiparisNo", comboBox3.Text);
                                    komut4.Parameters.AddWithValue("@PaketTarihi", tarih);
                                    komut4.Parameters.AddWithValue("@PaketSayısı", textBox11.Text);
                                    komut4.Parameters.AddWithValue("@Aşama", "Hazır");
                                    komut4.ExecuteNonQuery();
                                    label19.Text = comboBox3.Text;
                                    label14.Text = dataGridView1.Rows[0].Cells["Müşteri"].Value.ToString();
                                    liste();
                                    Methodlar();
                                }
                                textBox11.Text = "";
                                comboBox3.Text = "";
                            }

                        }
                        else
                        {
                            izin = "var";
                        }
                    }
                }

                if (izin == "var")
                {
                    if (radioButton1.Checked == true)
                    {
                        int j = 2;
                        Excel.Application excel = new Excel.Application();
                        excel.Visible = true;
                        object Missing = Type.Missing;
                        //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Modeks_Dosyalar\\KARGO.xlsx ");
                        //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Enes\\Desktop\\KARGO.xlsx");
                        Excel.Workbook workbook = excel.Workbooks.Open("C:\\Modeks_Dosyalar\\KARGO.xlsx");
                        Excel.Worksheet sheet1 = (Excel.Worksheet)workbook.Sheets[1];
                        sheet1.Cells[4, 1].Value = dataGridView1.Rows[0].Cells[0].Value.ToString();
                        sheet1.Cells[6, 1].Value = dataGridView1.Rows[0].Cells[1].Value.ToString();
                        sheet1.Cells[12, 1].Value = dataGridView1.Rows[0].Cells[7].Value.ToString();
                        sheet1.Cells[13, 1].Value = dataGridView1.Rows[0].Cells[2].Value.ToString();
                        sheet1.Cells[14, 1].Value = dataGridView1.Rows[0].Cells[8].Value.ToString();
                        sheet1.Cells[15, 1].Value = dataGridView1.Rows[0].Cells[9].Value.ToString();
                        sheet1.Cells[14, 4].Value = 1 + "/" + textBox11.Text;
                        sheet1.Cells[8, 1].Value = dataGridView1.Rows[0].Cells[11].Value.ToString();
                        sheet1.Cells[4, 3].Value = dataGridView1.Rows[0].Cells[10].Value.ToString();
                        for (int i = 0; i < Convert.ToInt32(textBox11.Text) - 1; i++)
                        {
                            sheet1.Range["A1:D15"].Copy(sheet1.Range["A" + (j + 14) + ""]);
                            j += 15;
                            sheet1.Cells[(12 + j), 4].Value = (i + 2) + "/" + textBox11.Text;
                            
                        }
                        //sheet1.Range["A14:F23"].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                        string etiketFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Etiket");

                        if (!Directory.Exists(etiketFolderPath))
                        {
                            Directory.CreateDirectory(etiketFolderPath);
                        }

                        string excelFilePath = Path.Combine(etiketFolderPath, "Paket" + dataGridView1.Rows[0].Cells[0].Value.ToString() + ".xlsx");


                        workbook.SaveAs(excelFilePath);
                        sheet1.PrintOutEx(Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing);
                        workbook.Close(false, Missing, Missing);
                        excel.Quit();

                        string tarih = Convert.ToDateTime(DateTime.Now).ToString("yyyy-MM-dd HH:mm:ss");

                        string sorgu4 = "UPDATE Siparişler SET PaketTarihi=@PaketTarihi, PaketSayısı=@PaketSayısı, Aşama=@Aşama WHERE SiparisNo=@SiparisNo AND AnaSiparişMi='" + "Evet" + "' ";
                        SqlCommand komut4;
                        komut4 = new SqlCommand(sorgu4, bgl.baglanti());
                        komut4.Parameters.AddWithValue("@SiparisNo", comboBox3.Text);
                        komut4.Parameters.AddWithValue("@PaketTarihi", tarih);
                        komut4.Parameters.AddWithValue("@PaketSayısı", textBox11.Text);
                        komut4.Parameters.AddWithValue("@Aşama", "Hazır");
                        komut4.ExecuteNonQuery();
                        label19.Text = comboBox3.Text;
                        label14.Text = dataGridView1.Rows[0].Cells["Müşteri"].Value.ToString();
                        liste();
                        Methodlar();
                    }
                    else if (radioButton2.Checked == true)
                    {
                        int j = 2;
                        Excel.Application excel = new Excel.Application();
                        excel.Visible = true;
                        object Missing = Type.Missing;
                        //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Enes\\Desktop\\KARGO.xlsx");
                        //Excel.Workbook workbook = excel.Workbooks.Open("C:\\Users\\Public\\Modeks_Dosyalar\\KARGO.xlsx ");
                        Excel.Workbook workbook = excel.Workbooks.Open("C:\\Modeks_Dosyalar\\KARGO.xlsx");
                        Excel.Worksheet sheet1 = (Excel.Worksheet)workbook.Sheets[2];
                        sheet1.Cells[4, 1].Value = dataGridView1.Rows[0].Cells[0].Value.ToString();
                        sheet1.Cells[6, 1].Value = dataGridView1.Rows[0].Cells[1].Value.ToString();
                        sheet1.Cells[12, 1].Value = dataGridView1.Rows[0].Cells[7].Value.ToString();
                        sheet1.Cells[13, 1].Value = dataGridView1.Rows[0].Cells[2].Value.ToString();
                        sheet1.Cells[14, 1].Value = dataGridView1.Rows[0].Cells[8].Value.ToString();
                        sheet1.Cells[15, 1].Value = dataGridView1.Rows[0].Cells[9].Value.ToString();
                        sheet1.Cells[14, 4].Value = 1 + "/" + textBox11.Text;
                        sheet1.Cells[8, 1].Value = dataGridView1.Rows[0].Cells[11].Value.ToString();
                        sheet1.Cells[4, 3].Value = dataGridView1.Rows[0].Cells[10].Value.ToString();
                        for (int i = 0; i < Convert.ToInt32(textBox11.Text) - 1; i++)
                        {
                            sheet1.Range["A1:D15"].Copy(sheet1.Range["A" + (j + 14) + ""]);
                            j += 15;
                            sheet1.Cells[(12 + j), 4].Value = (i + 2) + "/" + textBox11.Text;

                        }
                        //sheet1.Range["A14:F23"].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                        //sheet1.PrintOutEx(Missing, Missing, Missing, Missing, Missing, Missing, Missing, Missing);
                        string etiketFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Etiket");

                        if (!Directory.Exists(etiketFolderPath))
                        {
                            Directory.CreateDirectory(etiketFolderPath);
                        }

                        string excelFilePath = Path.Combine(etiketFolderPath, "Paket" + dataGridView1.Rows[0].Cells[0].Value.ToString() + ".xlsx");


                        workbook.SaveAs(excelFilePath);
                        workbook.Close(false, Missing, Missing);
                        excel.Quit();

                        string tarih = Convert.ToDateTime(DateTime.Now).ToString("yyyy-MM-dd HH:mm:ss");

                        string sorgu4 = "UPDATE Siparişler SET PaketTarihi=@PaketTarihi, PaketSayısı=@PaketSayısı, Aşama=@Aşama WHERE SiparisNo=@SiparisNo AND AnaSiparişMi='" + "Evet" + "' ";
                        SqlCommand komut4;
                        komut4 = new SqlCommand(sorgu4, bgl.baglanti());
                        komut4.Parameters.AddWithValue("@SiparisNo", comboBox3.Text);
                        komut4.Parameters.AddWithValue("@PaketTarihi", tarih);
                        komut4.Parameters.AddWithValue("@PaketSayısı", textBox11.Text);
                        komut4.Parameters.AddWithValue("@Aşama", "Hazır");
                        komut4.ExecuteNonQuery();
                        label19.Text = comboBox3.Text;
                        label14.Text = dataGridView1.Rows[0].Cells["Müşteri"].Value.ToString();
                        liste();
                        Methodlar();
                    }

                    textBox11.Text = "";
                    comboBox3.Text = "";

                }
                else
                {
                    MessageBox.Show("Lütfen geçerli bir sipariş numarası giriniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Kaç paket yaptıysanız onu yazınız.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            liste();
            Methodlar();
            textBox11.Text = "";
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            siparisnoyagöresırala();
            aynısipnugetirme();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            müşteriyegöresırala();
            aynısipnugetirme();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            liste();
            aynısipnugetirme();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Form9 frm = new Form9();
            this.Hide();
            frm.yetki = yetki;
            frm.Show();
        }
        string siparisnoacil, renkacil;
        private void TeslimTarihine3GünKalanlarıYakSöndür()
        {
            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT * FROM Siparişler WHERE CONVERT(datetime, TeslimTarihi, 104) <= DATEADD(day, 6, GETDATE()) AND Onay=@Onay AND AnaSiparişMi=@p1 AND PaketTarihi is null ORDER BY SiparisNo ASC";
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

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void Form9_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Hide();
            Form1 form1 = Application.OpenForms["Form1"] as Form1;
            if (form1 != null)
            {
                form1.Show();
            }
        }

        private void comboBox3_TextChanged(object sender, EventArgs e)
        {
            string srg = comboBox3.Text;
            string sorgu = "SELECT DISTINCT SiparisNo,Müşteri,Renk,SiparişTipi,ToplamM2,ToplamAdet,SevkTürü,Model,SiparişTarihi,TeslimTarihi,BID,Adres,PaketTarihi From Siparişler where SiparisNo Like '" + srg + "' AND AnaSiparişMi = '" + "Evet" + "' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            aynısipnugetirme();
            onayrenklendir();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        private void bugünegöresırala()
        {
            DateTime bitir = DateTime.Now;
            DateTime basla = DateTime.Now;
            dateTimePicker1.Value = basla;
            label11.Text = basla.ToString("yyyy - MM - dd");
            label13.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");
            string sorgu = "SELECT DISTINCT SiparisNo,Müşteri,Renk,SiparişTipi,ToplamM2,ToplamAdet,SevkTürü,Model,SiparişTarihi,TeslimTarihi,BID,Adres,PaketTarihi From Siparişler where SiparişTarihi between '" + label11.Text + "' AND '" + label13.Text + "' AND AnaSiparişMi='Evet' AND Aşama='Kargo' ORDER BY SiparisNo ASC";
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
            label11.Text = basla.ToString("yyyy - MM - dd HH:mm:ss");
            label13.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");
            string sorgu = "SELECT DISTINCT SiparisNo,Müşteri,Renk,SiparişTipi,ToplamM2,ToplamAdet,SevkTürü,Model,SiparişTarihi,TeslimTarihi,BID,Adres,PaketTarihi From Siparişler where SiparişTarihi between '" + label11.Text + "' AND '" + label13.Text + "' AND AnaSiparişMi='Evet' AND Aşama='Kargo' ORDER BY SiparisNo ASC";
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
            label11.Text = basla.ToString("yyyy - MM - dd HH:mm:ss");
            label13.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");
            string sorgu = "SELECT DISTINCT SiparisNo,Müşteri,Renk,SiparişTipi,ToplamM2,ToplamAdet,SevkTürü,Model,SiparişTarihi,TeslimTarihi,BID,Adres,PaketTarihi From Siparişler where SiparişTarihi between '" + label11.Text + "' AND '" + label13.Text + "' AND AnaSiparişMi='Evet' AND Aşama='Kargo' ORDER BY SiparisNo ASC";
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
            label11.Text = basla.ToString("yyyy - MM - dd HH:mm:ss");
            label13.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");
            string sorgu = "SELECT DISTINCT SiparisNo,Müşteri,Renk,SiparişTipi,ToplamM2,ToplamAdet,SevkTürü,Model,SiparişTarihi,TeslimTarihi,BID,Adres,PaketTarihi From Siparişler where SiparişTarihi between '" + label11.Text + "' AND '" + label13.Text + "' AND AnaSiparişMi='Evet' AND Aşama='Kargo' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            onayrenklendir();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            bugünegöresırala();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            haftayagöresırala();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ayagöresırala();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            yılagöresırala();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTime bitir = dateTimePicker1.Value;
            DateTime basla = dateTimePicker2.Value;
            label11.Text = basla.ToString("yyyy - MM - dd");
            label13.Text = bitir.ToString("yyyy - MM - dd");
            string sorgu = "SELECT DISTINCT SiparisNo,Müşteri,Renk,SiparişTipi,ToplamM2,ToplamAdet,SevkTürü,Model,SiparişTarihi,TeslimTarihi,BID,Adres,PaketTarihi From Siparişler where SiparişTarihi between '" + label11.Text + "' AND '" + label13.Text + "' AND AnaSiparişMi='Evet' AND Aşama='Kargo' ORDER BY SiparisNo ASC";
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
            label11.Text = basla.ToString("yyyy - MM - dd");
            label13.Text = bitir.ToString("yyyy - MM - dd");
            string sorgu = "SELECT DISTINCT SiparisNo,Müşteri,Renk,SiparişTipi,ToplamM2,ToplamAdet,SevkTürü,Model,SiparişTarihi,TeslimTarihi,BID,Adres,PaketTarihi From Siparişler where SiparişTarihi between '" + label11.Text + "' AND '" + label13.Text + "' AND AnaSiparişMi='Evet' AND Aşama='Kargo' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            onayrenklendir();
        }

        private void dataGridView1_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {

        }

        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
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
        }

        private void button27_Click(object sender, EventArgs e)
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
                frm.yetki = yetki;
                frm.hangiformdan = "Form7";
                this.Hide();
                frm.ShowDialog();
            }
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
