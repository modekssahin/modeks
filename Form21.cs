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
    public partial class Form21 : Form
    {
        public Form21()
        {
            InitializeComponent();
        }
        sqlsinif bgl = new sqlsinif();
        public string yetki;
        public string kullaniciadi;
        private void timer1_Tick(object sender, EventArgs e)
        {
            label2.Text = DateTime.Now.ToLongDateString();
            label12.Text = DateTime.Now.ToLongTimeString();
        }

        private void Form21_Load(object sender, EventArgs e)
        {
            DateTime bitir = DateTime.Now;
            DateTime basla = DateTime.Now;
            dateTimePicker1.Value = basla;
            label21.Text = basla.ToString("yyyy - MM - dd");
            label22.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");
            timer1.Start();
            Methodlar();
        }
        private void Methodlar()
        {
            onaylıtümsiparişler();
            onaylanmayantümsiparişler();
            cnckesilen();
            etiketvepaletolan();
            etiketvepaletbekleyen();
            membranpresbekleyen();
            membranpresbasılan();
            paketkargoyapılan();
            paketkargobekleyen();
            teslimedilen();
            teslimbekleyen();
            cnckesimbekleyen();
            teslimedilenfabrikakargo();
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
            Methodlar();
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            DateTime bitir = dateTimePicker2.Value;
            DateTime basla = dateTimePicker1.Value;
            label21.Text = basla.ToString("yyyy - MM - dd");
            label22.Text = bitir.ToString("yyyy - MM - dd  HH:mm:ss");
            Methodlar();
        }

        private void button20_Click(object sender, EventArgs e)
        {
            DateTime bitir = DateTime.Now;
            DateTime basla = DateTime.Now;
            dateTimePicker1.Value = basla;
            label21.Text = basla.ToString("yyyy - MM - dd");
            label22.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DateTime bitir = DateTime.Now;
            DateTime bugun = DateTime.Today;
            DateTime basla = bugun.AddDays(-(int)bugun.DayOfWeek + 1);
            if (bugun.DayOfWeek.ToString()=="Sunday")
            {
                basla = basla.AddDays(-7);
            }
            dateTimePicker1.Value = basla;
            label21.Text = basla.ToString("yyyy - MM - dd HH:mm:ss");
            label22.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DateTime bitir = DateTime.Now;
            DateTime bugun = DateTime.Today;
            DateTime basla = new DateTime(bugun.Year, bugun.Month, 1);
            dateTimePicker1.Value = basla;
            label21.Text = basla.ToString("yyyy - MM - dd HH:mm:ss");
            label22.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DateTime bitir = DateTime.Now;
            DateTime bugun = DateTime.Today;
            DateTime basla = new DateTime(bugun.Year, 1, 1);
            dateTimePicker1.Value = basla;
            label21.Text = basla.ToString("yyyy - MM - dd HH:mm:ss");
            label22.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");
        }

        double onaylıtümsiparislerm2;
        double onaylanmayantümsiparişlerm2;
        double cnckesimbekleyenm2;
        double cnckesilenm2;
        double etiketvepaletolanm2;
        double etiketvepaletbekleyenm2;
        double membranpresbekleyenm2;
        double membranpresbasılanm2;
        double paketkargoyapılanm2;
        double paketkargobekleyenm2;
        double teslimbekleyenm2;
        double teslimedilenm2;
        double teslimedilenkargom2;
        double teslimedilenfabrikam2;
        private void onaylıtümsiparişler()
        {
            try
            {
                onaylıtümsiparislerm2 = 0;
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Siparişler where OnayTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi=@p1 AND Onay=@Onay";
                komut.Parameters.AddWithValue("@p1", "Evet");
                komut.Parameters.AddWithValue("@Onay", "Onaylandı");
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    onaylıtümsiparislerm2 += Convert.ToDouble(dr["M2"].ToString());
                }
                button6.Text = onaylıtümsiparislerm2.ToString("0.##") + " m²";
            }
            catch (Exception)
            {

            }
           
        }
        private void cnckesimbekleyen()
        {
            try
            {
                cnckesimbekleyenm2 = 0;
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Siparişler where OnayTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi=@p1 AND Onay=@Onay AND KesildiMi=@KesildiMi";
                komut.Parameters.AddWithValue("@p1", "Evet");
                komut.Parameters.AddWithValue("@Onay", "Onaylandı");
                komut.Parameters.AddWithValue("@KesildiMi", "Hayır");
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    cnckesimbekleyenm2 += Convert.ToDouble(dr["M2"].ToString());
                }
                button1.Text = cnckesimbekleyenm2.ToString("0.##") + " m²";
            }
            catch (Exception)
            {

            }

        }
        private void onaylanmayantümsiparişler()
        {
            try
            {
                onaylanmayantümsiparişlerm2 = 0;
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Siparişler where SiparişTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi=@p1 AND Onay=@Onay";
                komut.Parameters.AddWithValue("@p1", "Evet");
                komut.Parameters.AddWithValue("@Onay", "Onay Bekliyor");
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    onaylanmayantümsiparişlerm2 += Convert.ToDouble(dr["M2"].ToString());
                }
                button17.Text = onaylanmayantümsiparişlerm2.ToString("0.##") + " m²";
            }
            catch (Exception)
            {

            }

        }
        private void cnckesilen()
        {
            try
            {
                cnckesilenm2 = 0;
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Siparişler where KesildiTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi=@p1 AND Onay=@Onay AND KesildiMi=@KesildiMi";
                komut.Parameters.AddWithValue("@p1", "Evet");
                komut.Parameters.AddWithValue("@Onay", "Onaylandı");
                komut.Parameters.AddWithValue("@KesildiMi", "Evet");
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    cnckesilenm2 += Convert.ToDouble(dr["M2"].ToString());
                }
                button16.Text = cnckesilenm2.ToString("0.##") + " m²";
            }
            catch (Exception)
            {

            }

        }
        private void etiketvepaletbekleyen()
        {
            try
            {
                etiketvepaletbekleyenm2 = 0;
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Siparişler where KesildiTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi=@p1 AND Onay=@Onay AND KesildiMi=@KesildiMi AND Etiket is null";
                komut.Parameters.AddWithValue("@p1", "Evet");
                komut.Parameters.AddWithValue("@Onay", "Onaylandı");
                komut.Parameters.AddWithValue("@KesildiMi", "Evet");
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    etiketvepaletbekleyenm2 += Convert.ToDouble(dr["M2"].ToString());
                }
                button5.Text = etiketvepaletbekleyenm2.ToString("0.##") + " m²";
            }
            catch (Exception)
            {

            }

        }
        private void etiketvepaletolan()
        {
            try
            {
                etiketvepaletolanm2 = 0;
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Siparişler where Etiket between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi=@p1 AND Onay=@Onay AND KesildiMi=@KesildiMi AND Etiket is not null";
                komut.Parameters.AddWithValue("@p1", "Evet");
                komut.Parameters.AddWithValue("@Onay", "Onaylandı");
                komut.Parameters.AddWithValue("@KesildiMi", "Evet");
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    etiketvepaletolanm2 += Convert.ToDouble(dr["M2"].ToString());
                }
                button15.Text = etiketvepaletolanm2.ToString("0.##") + " m²";
            }
            catch (Exception)
            {

            }

        }
        private void membranpresbekleyen()
        {
            try
            {
                membranpresbekleyenm2 = 0;
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Siparişler where Etiket between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi=@p1 AND Onay=@Onay AND KesildiMi=@KesildiMi AND Etiket is not null AND Aşama=@Aşama";
                komut.Parameters.AddWithValue("@p1", "Evet");
                komut.Parameters.AddWithValue("@Onay", "Onaylandı");
                komut.Parameters.AddWithValue("@KesildiMi", "Evet");
                komut.Parameters.AddWithValue("@Aşama", "Palet");
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    membranpresbekleyenm2 += Convert.ToDouble(dr["M2"].ToString());
                }
                button7.Text = membranpresbekleyenm2.ToString("0.##") + " m²";
            }
            catch (Exception)
            {

            }

        }
        private void membranpresbasılan()
        {
            try
            {
                membranpresbasılanm2 = 0;
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Siparişler where MembranPressTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi=@p1 AND Onay=@Onay AND KesildiMi=@KesildiMi AND Etiket is not null AND MembranPressTarihi is not null";
                komut.Parameters.AddWithValue("@p1", "Evet");
                komut.Parameters.AddWithValue("@Onay", "Onaylandı");
                komut.Parameters.AddWithValue("@KesildiMi", "Evet");
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    membranpresbasılanm2 += Convert.ToDouble(dr["M2"].ToString());
                }
                button14.Text = membranpresbasılanm2.ToString("0.##") + " m²";
            }
            catch (Exception)
            {

            }

        }
        private void paketkargobekleyen()
        {
            try
            {
                paketkargobekleyenm2 = 0;
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Siparişler where MembranPressTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi=@p1 AND Onay=@Onay AND KesildiMi=@KesildiMi AND Etiket is not null AND MembranPressTarihi is not null AND PaketTarihi is null";
                komut.Parameters.AddWithValue("@p1", "Evet");
                komut.Parameters.AddWithValue("@Onay", "Onaylandı");
                komut.Parameters.AddWithValue("@KesildiMi", "Evet");
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    paketkargobekleyenm2 += Convert.ToDouble(dr["M2"].ToString());
                }
                button8.Text = paketkargobekleyenm2.ToString("0.##") + " m²";
            }
            catch (Exception)
            {

            }

        }
        private void paketkargoyapılan()
        {
            try
            {
                paketkargoyapılanm2 = 0;
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Siparişler where PaketTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi=@p1 AND Onay=@Onay AND KesildiMi=@KesildiMi AND Etiket is not null AND MembranPressTarihi is not null AND PaketTarihi is not null";
                komut.Parameters.AddWithValue("@p1", "Evet");
                komut.Parameters.AddWithValue("@Onay", "Onaylandı");
                komut.Parameters.AddWithValue("@KesildiMi", "Evet");
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    paketkargoyapılanm2 += Convert.ToDouble(dr["M2"].ToString());
                }
                button13.Text = paketkargoyapılanm2.ToString("0.##") + " m²";
            }
            catch (Exception)
            {

            }

        }
        private void teslimbekleyen()
        {
            try
            {
                teslimbekleyenm2 = 0;
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Siparişler where PaketTarihi between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi=@p1 AND Onay=@Onay AND KesildiMi=@KesildiMi AND Etiket is not null AND MembranPressTarihi is not null AND PaketTarihi is not null AND TeslimEdilenTarih is null";
                komut.Parameters.AddWithValue("@p1", "Evet");
                komut.Parameters.AddWithValue("@Onay", "Onaylandı");
                komut.Parameters.AddWithValue("@KesildiMi", "Evet");
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    teslimbekleyenm2 += Convert.ToDouble(dr["M2"].ToString());
                }
                button9.Text = teslimbekleyenm2.ToString("0.##") + " m²";
            }
            catch (Exception)
            {

            }

        }
        private void teslimedilen()
        {
            try
            {
                teslimedilenm2 = 0;
                teslimedilenfabrikam2 = 0;
                teslimedilenkargom2 = 0;
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Siparişler where TeslimEdilenTarih between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi=@p1 AND Onay=@Onay AND KesildiMi=@KesildiMi AND Etiket is not null AND MembranPressTarihi is not null AND PaketTarihi is not null AND TeslimEdilenTarih is not null";
                komut.Parameters.AddWithValue("@p1", "Evet");
                komut.Parameters.AddWithValue("@Onay", "Onaylandı");
                komut.Parameters.AddWithValue("@KesildiMi", "Evet");
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    teslimedilenm2 += Convert.ToDouble(dr["M2"].ToString());
                }
                button12.Text = teslimedilenm2.ToString("0.##") + " m²";
            }
            catch (Exception)
            {

            }

        }
        private void teslimedilenfabrikakargo()
        {
            try
            {
                teslimedilenfabrikam2 = 0;
                teslimedilenkargom2 = 0;
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Siparişler where TeslimEdilenTarih between '" + label21.Text + "' AND '" + label22.Text + "' AND AnaSiparişMi=@p1 AND Onay=@Onay AND KesildiMi=@KesildiMi AND Etiket is not null AND MembranPressTarihi is not null AND PaketTarihi is not null AND TeslimEdilenTarih is not null";
                komut.Parameters.AddWithValue("@p1", "Evet");
                komut.Parameters.AddWithValue("@Onay", "Onaylandı");
                komut.Parameters.AddWithValue("@KesildiMi", "Evet");
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    if (dr["SevkTürü"].ToString() == "Fabrika")
                    {
                        teslimedilenfabrikam2 += Convert.ToDouble(dr["M2"].ToString());
                    }
                    if (dr["SevkTürü"].ToString() == "Kargo")
                    {
                        teslimedilenkargom2 += Convert.ToDouble(dr["M2"].ToString());
                    }
                }
                button10.Text = teslimedilenfabrikam2.ToString("0.##") + " m²";
                button11.Text = teslimedilenkargom2.ToString("0.##") + " m²";
            }
            catch (Exception)
            {

            }

        }

    }
}
