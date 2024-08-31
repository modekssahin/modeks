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
    public partial class Form20 : Form
    {
        public Form20()
        {
            InitializeComponent();
        }
        sqlsinif bgl = new sqlsinif();
        public string yetki;
        public string kullaniciadi;
        private void liste()
        {
            string kayit = "SELECT SiparisNo,Müşteri, SiparişTipi,SiparişTarihi,Onay,SevkTürü,ToplamAdet,ToplamM2,ToplamFiyat,AcilFarkı,Kargo,Tasarım,M2KapakFarkı From Siparişler where AnaSiparişMi=@p1 ORDER BY SiparisNo ASC";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            komut.Parameters.AddWithValue("@p1", "Evet");
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            aynısipnugetirme();
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
        private void timer1_Tick(object sender, EventArgs e)
        {
            label2.Text = DateTime.Now.ToLongDateString();
            label12.Text = DateTime.Now.ToLongTimeString();
        }

        private void Form20_Load(object sender, EventArgs e)
        {
            timer1.Start();
            liste();
            DateTime bitir = DateTime.Now;
            DateTime basla = DateTime.Now;
            dateTimePicker1.Value = basla;
            label21.Text = basla.ToString("yyyy - MM - dd");
            label22.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");

            string sorgu = "SELECT SiparisNo,Müşteri, SiparişTipi,SiparişTarihi,Onay,SevkTürü,ToplamAdet,ToplamM2,ToplamFiyat,AcilFarkı,Kargo,Tasarım,M2KapakFarkı From Siparişler where SiparişTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            aynısipnugetirme();
            SatırlarınEnAltınaToplat();
        }

        private void pictureBox8_Click(object sender, EventArgs e)
        {
            Form17 frm = new Form17();
            frm.yetki = yetki;
            this.Hide();
            frm.ShowDialog();
        }

        private void groupBox1_Paint(object sender, PaintEventArgs e)
        {
            GroupBox box = sender as GroupBox;
            DrawGroupBox(box, e.Graphics, Color.Black, Color.Black);
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTime bitir = dateTimePicker2.Value;
            DateTime basla = dateTimePicker1.Value;
            label21.Text = basla.ToString("yyyy - MM - dd");
            label22.Text = bitir.ToString("yyyy - MM - dd  HH:mm:ss");

            string sorgu = "SELECT SiparisNo,Müşteri, SiparişTipi,SiparişTarihi,Onay,SevkTürü,ToplamAdet,ToplamM2,ToplamFiyat,AcilFarkı,Kargo,Tasarım,M2KapakFarkı From Siparişler where SiparişTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            aynısipnugetirme();
            SatırlarınEnAltınaToplat();
        }

        private void button20_Click(object sender, EventArgs e)
        {
            button20.BackColor = Color.Green;
            button2.BackColor = Color.LightGray;
            button3.BackColor = Color.LightGray;
            button4.BackColor = Color.LightGray;
            DateTime bitir = DateTime.Now;
            DateTime basla = DateTime.Now;
            dateTimePicker1.Value = basla;
            label21.Text = basla.ToString("yyyy - MM - dd");
            label22.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");

            string sorgu = "SELECT SiparisNo,Müşteri, SiparişTipi,SiparişTarihi,Onay,SevkTürü,ToplamAdet,ToplamM2,ToplamFiyat,AcilFarkı,Kargo,Tasarım,M2KapakFarkı From Siparişler where SiparişTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            aynısipnugetirme();
            SatırlarınEnAltınaToplat();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button20.BackColor = Color.LightGray;
            button2.BackColor = Color.Green;
            button3.BackColor = Color.LightGray;
            button4.BackColor = Color.LightGray;
            DateTime bitir = DateTime.Now;
            DateTime bugun = DateTime.Today;
            DateTime basla = bugun.AddDays(-(int)bugun.DayOfWeek + 1);
            dateTimePicker1.Value = basla;
            label21.Text = basla.ToString("yyyy - MM - dd HH:mm:ss");
            label22.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");

            string sorgu = "SELECT SiparisNo,Müşteri, SiparişTipi,SiparişTarihi,Onay,SevkTürü,ToplamAdet,ToplamM2,ToplamFiyat,AcilFarkı,Kargo,Tasarım,M2KapakFarkı From Siparişler where SiparişTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            aynısipnugetirme();
            SatırlarınEnAltınaToplat();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            button20.BackColor = Color.LightGray;
            button2.BackColor = Color.LightGray;
            button3.BackColor = Color.Green;
            button4.BackColor = Color.LightGray;
            DateTime bitir = DateTime.Now;
            DateTime bugun = DateTime.Today;
            DateTime basla = new DateTime(bugun.Year, bugun.Month, 1);
            dateTimePicker1.Value = basla;
            label21.Text = basla.ToString("yyyy - MM - dd HH:mm:ss");
            label22.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");

            string sorgu = "SELECT SiparisNo,Müşteri, SiparişTipi,SiparişTarihi,Onay,SevkTürü,ToplamAdet,ToplamM2,ToplamFiyat,AcilFarkı,Kargo,Tasarım,M2KapakFarkı From Siparişler where SiparişTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            aynısipnugetirme();
            SatırlarınEnAltınaToplat();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            button20.BackColor = Color.LightGray;
            button2.BackColor = Color.LightGray;
            button3.BackColor = Color.LightGray;
            button4.BackColor = Color.Green;
            DateTime bitir = DateTime.Now;
            DateTime bugun = DateTime.Today;
            DateTime basla = new DateTime(bugun.Year, 1, 1);
            dateTimePicker1.Value = basla;
            label21.Text = basla.ToString("yyyy - MM - dd HH:mm:ss");
            label22.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");

            string sorgu = "SELECT SiparisNo,Müşteri, SiparişTipi,SiparişTarihi,Onay,SevkTürü,ToplamAdet,ToplamM2,ToplamFiyat,AcilFarkı,Kargo,Tasarım,M2KapakFarkı From Siparişler where SiparişTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            aynısipnugetirme();
            SatırlarınEnAltınaToplat();
        }
        private void onaylanmışsiparişler()
        {
            string srg = "Onaylandı";
            string sorgu = "SELECT SiparisNo,Müşteri, SiparişTipi,OnayTarihi,Onay,SevkTürü,ToplamAdet,ToplamM2,ToplamFiyat,AcilFarkı,Kargo,Tasarım,M2KapakFarkı From Siparişler where Onay Like '" + srg + "' AND OnayTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            aynısipnugetirme();
        }
        private void cnc()
        {
            string srg = "Onaylandı";
            string sorgu = "SELECT SiparisNo,Müşteri, SiparişTipi,KesildiTarihi,Onay,SevkTürü,ToplamAdet,ToplamM2,ToplamFiyat,AcilFarkı,Kargo,Tasarım,M2KapakFarkı From Siparişler where Onay Like '" + srg + "' AND KesildiTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            aynısipnugetirme();
        }
        private void etiket()
        {
            string srg = "Onaylandı";
            string sorgu = "SELECT SiparisNo,Müşteri,SiparişTipi,Etiket,Onay,SevkTürü,ToplamAdet,ToplamM2,ToplamFiyat,AcilFarkı,Kargo,Tasarım,M2KapakFarkı From Siparişler where Onay Like '" + srg + "' AND Etiket between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            aynısipnugetirme();
        }
        private void press()
        {
            string srg = "Onaylandı";
            string sorgu = "SELECT SiparisNo,Müşteri, SiparişTipi,MembranPressTarihi,Onay,SevkTürü,ToplamAdet,ToplamM2,ToplamFiyat,AcilFarkı,Kargo,Tasarım,M2KapakFarkı From Siparişler where Onay Like '" + srg + "' AND MembranPressTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            aynısipnugetirme();
        }
        private void paket()
        {
            string srg = "Onaylandı";
            string sorgu = "SELECT SiparisNo,Müşteri, SiparişTipi,PaketTarihi,Onay,SevkTürü,ToplamAdet,ToplamM2,ToplamFiyat,AcilFarkı,Kargo,Tasarım,M2KapakFarkı From Siparişler where Onay Like '" + srg + "' AND PaketTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            aynısipnugetirme();
        }
        private void teslim()
        {
            string srg = "Onaylandı";
            string sorgu = "SELECT SiparisNo,Müşteri, SiparişTipi,TeslimEdilenTarih,Onay,SevkTürü,ToplamAdet,ToplamM2,ToplamFiyat,AcilFarkı,Kargo,Tasarım,M2KapakFarkı From Siparişler where Onay Like '" + srg + "' AND TeslimEdilenTarih between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            aynısipnugetirme();
        }
        private void fabrika()
        {
            string srg = "Fabrika";
            string sorgu = "SELECT SiparisNo,Müşteri, SiparişTipi,TeslimEdilenTarih,Onay,SevkTürü,ToplamAdet,ToplamM2,ToplamFiyat,AcilFarkı,Kargo,Tasarım,M2KapakFarkı From Siparişler where SevkTürü Like '" + srg + "' AND TeslimEdilenTarih between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            aynısipnugetirme();
        }
        private void kargo()
        {
            string srg = "Kargo";
            string sorgu = "SELECT SiparisNo,Müşteri, SiparişTipi,TeslimEdilenTarih,Onay,SevkTürü,ToplamAdet,ToplamM2,ToplamFiyat,AcilFarkı,Kargo,Tasarım,M2KapakFarkı From Siparişler where SevkTürü Like '" + srg + "' AND TeslimEdilenTarih between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            aynısipnugetirme();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            label3.Text = "ONAYLANANLAR";
            onaylanmışsiparişler();
            SatırlarınEnAltınaToplat();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            label3.Text = "CNC";
            cnc();
            SatırlarınEnAltınaToplat();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            label3.Text = "ETİKET";
            etiket();
            SatırlarınEnAltınaToplat();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            label3.Text = "MEMBRAN/PRESS";
            press();
            SatırlarınEnAltınaToplat();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            label3.Text = "PAKET/KARGO";
            paket();
            SatırlarınEnAltınaToplat();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            label3.Text = "TESLİM";
            teslim();
            SatırlarınEnAltınaToplat();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            label3.Text = "FABRİKA";
            fabrika();
            SatırlarınEnAltınaToplat();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            label3.Text = "KARGO";
            kargo();
            SatırlarınEnAltınaToplat();
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            DateTime bitir = dateTimePicker2.Value;
            DateTime basla = dateTimePicker1.Value;
            label21.Text = basla.ToString("yyyy - MM - dd");
            label22.Text = bitir.ToString("yyyy - MM - dd  HH:mm:ss");

            string sorgu = "SELECT SiparisNo,Müşteri, SiparişTipi,SiparişTarihi,Onay,SevkTürü,ToplamAdet,ToplamM2,ToplamFiyat,AcilFarkı,Kargo,Tasarım,M2KapakFarkı From Siparişler where SiparişTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi='" + "Evet" + "'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
            aynısipnugetirme();
            SatırlarınEnAltınaToplat();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Form17 frm = new Form17();
            frm.yetki = yetki;
            this.Hide();
            frm.ShowDialog();
        }

        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            aynısipnugetirme();
        }
        double satırlarınM2Toplamı = 0;
        double satırlarınAdetToplamı = 0;
        double satırlarınFiyatToplamı = 0;
        double satırlarınAcilFarkToplamı = 0;
        double satırlarınM2KapakFarkıToplamı = 0;
        double satırlarınKargoToplamı = 0;
        private void SatırlarınEnAltınaToplat()
        {
            satırlarınM2Toplamı = 0;
            satırlarınAdetToplamı = 0;
            satırlarınFiyatToplamı = 0;
            satırlarınAcilFarkToplamı = 0;
            satırlarınKargoToplamı = 0;
            satırlarınM2KapakFarkıToplamı = 0;
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                if (dataGridView1.Rows[i].Visible == true)
                {
                    satırlarınM2Toplamı += Convert.ToDouble(dataGridView1.Rows[i].Cells["ToplamM2"].Value);
                    satırlarınAdetToplamı += Convert.ToDouble(dataGridView1.Rows[i].Cells["ToplamAdet"].Value);
                    satırlarınFiyatToplamı += Convert.ToDouble(dataGridView1.Rows[i].Cells["ToplamFiyat"].Value);
                    satırlarınAcilFarkToplamı += Convert.ToDouble(dataGridView1.Rows[i].Cells["AcilFarkı"].Value);
                    satırlarınKargoToplamı += Convert.ToDouble(dataGridView1.Rows[i].Cells["Kargo"].Value);
                    satırlarınM2KapakFarkıToplamı += Convert.ToDouble(dataGridView1.Rows[i].Cells["M2KapakFarkı"].Value);
                }
            }
            dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells["ToplamM2"].Value = satırlarınM2Toplamı.ToString("#,##0.00");
            dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells["ToplamAdet"].Value = satırlarınAdetToplamı.ToString("#,##0.00");
            dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells["ToplamFiyat"].Value = satırlarınFiyatToplamı.ToString("#,##0.00");
            dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells["Kargo"].Value = satırlarınKargoToplamı.ToString("#,##0.00");
            dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells["AcilFarkı"].Value = satırlarınAcilFarkToplamı.ToString("#,##0.00");
            dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells["M2KapakFarkı"].Value = satırlarınM2KapakFarkıToplamı.ToString("#,##0.00");
            dataGridView1.Rows[dataGridView1.Rows.Count - 1].DefaultCellStyle.Font = new Font("Roboto", 10);
            dataGridView1.Rows[dataGridView1.Rows.Count - 1].DefaultCellStyle.ForeColor = Color.Red;


        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_Click(object sender, EventArgs e)
        {
            SatırlarınEnAltınaToplat();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
