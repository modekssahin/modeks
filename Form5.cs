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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Modeks
{
    public partial class Form5 : Form
    {
        public Form5()
        {
            InitializeComponent();
        }
        sqlsinif bgl = new sqlsinif();
        public string kullaniciadi;
        public string yetki;
        public string hangiformdan;
        private void liste()
        {
            string kayit = @"SELECT 
*,  
CASE WHEN (ISNULL(TRY_CAST(REPLACE(Kalan, ',', '.') AS decimal(12, 2)), 0) > 200) THEN('Stokta Yeterince Var')
WHEN (ISNULL(TRY_CAST(REPLACE(Kalan, ',', '.') AS decimal(12, 2)), 0) > 50) THEN('Sipariş Ver')
WHEN (ISNULL(TRY_CAST(REPLACE(Kalan, ',', '.') AS decimal(12, 2)), 0) < 50) THEN('Acil Sipariş Verilecek')
END AS Renklendirme
FROM Stoklar ORDER BY Renklendirme asc";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;

            dataGridView1.RowPrePaint += new DataGridViewRowPrePaintEventHandler(dataGridView1_RowPrePaint);
        }

        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            DataGridViewRow row = dataGridView1.Rows[e.RowIndex];
            string renklendirme = row.Cells["Renklendirme"].Value?.ToString();

            if (renklendirme == "Stokta Yeterince Var")
            {
                row.DefaultCellStyle.BackColor = Color.Green;
                row.DefaultCellStyle.ForeColor = Color.White;
            }
            else if (renklendirme == "Sipariş Ver")
            {
                row.DefaultCellStyle.BackColor = Color.Yellow;
                row.DefaultCellStyle.ForeColor = Color.Black;
            }
            else if (renklendirme == "Acil Sipariş Verilecek")
            {
                row.DefaultCellStyle.BackColor = Color.Red;
                row.DefaultCellStyle.ForeColor = Color.White;
            }
        }
        private void liste2()
        {
            string kayit = "SELECT * from  StokGirişi";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView2.DataSource = dt;
        }
        private void malzemebilgigetir()
        {
            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT *FROM Stoklar";
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                comboBox1.Items.Add(dr["Malzeme"]);
                comboBox2.Items.Add(dr["Malzeme"]);
            }
        }
        private void stokadınagöresırala()
        {

            string srg = textBox1.Text;
            string sorgu = "SELECT * from Stoklar where Malzeme Like '%" + srg + "%'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Stoklar");
            this.dataGridView1.DataSource = ds.Tables[0];
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
        private void Form5_Load(object sender, EventArgs e)
        {
            timer1.Start();
            liste();
            liste2();
            malzemebilgigetir();
            dataGridView1.Columns["id"].Visible = false;
            dataGridView1.Focus();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            label2.Text = DateTime.Now.ToLongDateString();
            label12.Text = DateTime.Now.ToLongTimeString();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            textBox3.Text = DateTime.Now.ToShortDateString();
            comboBox1.SelectedItem = dataGridView1.CurrentRow.Cells["Malzeme"].Value.ToString();
            textBox5.Text = dataGridView1.CurrentRow.Cells["Cins"].Value.ToString();
            textBox2.Text = dataGridView1.CurrentRow.Cells["Birim"].Value.ToString();

            textBox9.Text = DateTime.Now.ToShortDateString();
            comboBox2.SelectedItem = dataGridView1.CurrentRow.Cells["Malzeme"].Value.ToString();
            textBox8.Text = dataGridView1.CurrentRow.Cells["Cins"].Value.ToString();
            textBox7.Text = dataGridView1.CurrentRow.Cells["MevcutStok"].Value.ToString();
            textBox4.Text = dataGridView1.CurrentRow.Cells["Birim"].Value.ToString();


            textBox10.Text = dataGridView1.CurrentRow.Cells["Malzeme"].Value.ToString();
            textBox11.Text = dataGridView1.CurrentRow.Cells["Cins"].Value.ToString();



        }
        double stok;
        double kalan;
        private void button1_Click(object sender, EventArgs e)
        {
            stok = 0;
            kalan = 0;
            if (textBox3.Text != "" && comboBox1.Text != "" && textBox5.Text != "" && textBox6.Text != "" && textBox2.Text != "")
            {
                string kayit = "insert into StokGirişi(GirişTarihi,Malzeme,Cins,Miktar,Birim)values (@p1,@p2,@p3,@p4,@p5)";
                SqlCommand cmd = new SqlCommand(kayit, bgl.baglanti());
                cmd.Parameters.AddWithValue("@p1", textBox3.Text);
                cmd.Parameters.AddWithValue("@p2", comboBox1.Text);
                cmd.Parameters.AddWithValue("@p3", textBox5.Text);
                cmd.Parameters.AddWithValue("@p4", textBox6.Text);
                cmd.Parameters.AddWithValue("@p5", textBox2.Text);
                cmd.ExecuteNonQuery();


                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT * FROM Stoklar where Malzeme=@Malzeme";
                komut.Parameters.AddWithValue("@Malzeme", comboBox1.Text);
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    stok = Convert.ToDouble(dr["MevcutStok"].ToString());
                    kalan = Convert.ToDouble(dr["Kalan"].ToString());
                }


                string sorgu2 = "UPDATE Stoklar SET MevcutStok=@p1, Kalan=@p2 WHERE Malzeme=@Malzeme";
                SqlCommand cmd2;
                cmd2 = new SqlCommand(sorgu2, bgl.baglanti());
                cmd2.Parameters.AddWithValue("@Malzeme", comboBox1.Text);
                cmd2.Parameters.AddWithValue("@p1", (stok + Convert.ToDouble(textBox6.Text)).ToString());
                cmd2.Parameters.AddWithValue("@p2", (kalan + Convert.ToDouble(textBox6.Text)).ToString());
                cmd2.ExecuteNonQuery();
                MessageBox.Show("Stok Ekleme İşlemi Gerçekleşti.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                liste();
                liste2();
            }
            else
            {
                MessageBox.Show("Stok ekleyebilmek için tüm alanları doldurunuz!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        double mevcutstok;
        double kullanılan;
        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox9.Text != "" && comboBox2.Text != "" && textBox8.Text != "" && textBox7.Text != "" && textBox4.Text != "")
            {

                string sorgu2 = "UPDATE Stoklar SET MevcutStok=@p1,Kalan=@p2, Cins=@p3 WHERE Malzeme=@Malzeme";
                SqlCommand cmd2;
                cmd2 = new SqlCommand(sorgu2, bgl.baglanti());
                cmd2.Parameters.AddWithValue("@Malzeme", comboBox2.Text);
                cmd2.Parameters.AddWithValue("@p1", textBox7.Text);
                cmd2.Parameters.AddWithValue("@p2", (mevcutstok - kullanılan).ToString());
                cmd2.Parameters.AddWithValue("@p3", textBox8.Text);
                cmd2.ExecuteNonQuery();
                MessageBox.Show("Stok Güncelleme İşlemi Gerçekleşti.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);

                mevcutstok = 0;
                kullanılan = 0;
                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT * FROM Stoklar where Malzeme=@Malzeme";
                komut.Parameters.AddWithValue("@Malzeme", comboBox2.Text);
                komut.Connection = bgl.baglanti();
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    mevcutstok = Convert.ToDouble(dr["MevcutStok"].ToString());
                    kullanılan = Convert.ToDouble(dr["Kullanılan"].ToString());
                }
                string sorgu3 = "UPDATE Stoklar SET Kalan=@p2 WHERE Malzeme=@Malzeme";
                SqlCommand cmd3;
                cmd3 = new SqlCommand(sorgu3, bgl.baglanti());
                cmd3.Parameters.AddWithValue("@Malzeme", comboBox2.Text);
                cmd3.Parameters.AddWithValue("@p2", (mevcutstok - kullanılan).ToString());
                cmd3.ExecuteNonQuery();
                liste();
                liste2();
            }
            else
            {
                MessageBox.Show("Stok güncelleyebilmek için tüm alanları doldurunuz!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                liste();
            }
            stokadınagöresırala();
        }

        private void Form5_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox10.Text != "")
            {
                //RENKLER
                string kayit = "insert into Renkler(RenkAdı,renkyolu)values (@p1,@p2)";
                SqlCommand cmd = new SqlCommand(kayit, bgl.baglanti());
                cmd.Parameters.AddWithValue("@p1", textBox10.Text);
                cmd.Parameters.AddWithValue("@p2", textBox12.Text);
                cmd.ExecuteNonQuery();

                //STOKLAR
                string kayitstoklar = "insert into Stoklar(Malzeme,MevcutStok,Kullanılan,Kalan,Birim,Cins,renkyolu)values (@p1,@p2,@p3,@p4,@p5,@p6,@p7)";
                SqlCommand cmdstoklar = new SqlCommand(kayitstoklar, bgl.baglanti());
                cmdstoklar.Parameters.AddWithValue("@p1", textBox10.Text);
                cmdstoklar.Parameters.AddWithValue("@p2", 0);
                cmdstoklar.Parameters.AddWithValue("@p3", 0);
                cmdstoklar.Parameters.AddWithValue("@p4", 0);
                cmdstoklar.Parameters.AddWithValue("@p5", "m²");
                cmdstoklar.Parameters.AddWithValue("@p6", textBox11.Text);
                cmdstoklar.Parameters.AddWithValue("@p7", textBox12.Text);
                cmdstoklar.ExecuteNonQuery();

                //GRAFİK
                string kayitgrafik = "insert into Grafik(Renk)values (@p1)";
                SqlCommand cmdgrafik = new SqlCommand(kayitgrafik, bgl.baglanti());
                cmdgrafik.Parameters.AddWithValue("@p1", textBox10.Text);
                cmdgrafik.ExecuteNonQuery();

                liste();
                liste2();
                MessageBox.Show("Renk başarıyla eklenmiştir!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                malzemebilgigetir();
            }
            else
            {
                MessageBox.Show("Renk ekleyebilmek için tüm alanları doldurunuz!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox10.Text != "")
            {
                DialogResult d1 = new DialogResult();
                d1 = MessageBox.Show("Silmek istediğinizden emin misiniz? ", "Uyarı", MessageBoxButtons.YesNo);
                if (d1 == DialogResult.Yes)
                {
                    string sorgu = "DELETE FROM Renkler WHERE RenkAdı=@RenkAdı";
                    SqlCommand komut;
                    komut = new SqlCommand(sorgu, bgl.baglanti());
                    komut.Parameters.AddWithValue("@RenkAdı", textBox10.Text);
                    komut.ExecuteNonQuery();

                    string sorgu2 = "DELETE FROM Stoklar WHERE Malzeme=@Malzeme";
                    SqlCommand komut2;
                    komut2 = new SqlCommand(sorgu2, bgl.baglanti());
                    komut2.Parameters.AddWithValue("@Malzeme", textBox10.Text);
                    komut2.ExecuteNonQuery();

                    MessageBox.Show("Renk başarıyla silindi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    liste();
                    liste2();
                    malzemebilgigetir();

                }
            }
            else
            {
                MessageBox.Show("Renk Silebilmek İçin listeden bir kayıt seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button21_Click(object sender, EventArgs e)
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
            else
            {
                Form2 frm = new Form2();
                frm.yetki = yetki;
                frm.kullaniciadi = kullaniciadi;
                this.Hide();
                frm.Show();
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "Resim Formatı |*.png; *.jpg; *.pjeg";
            file.ShowDialog();
            string tamYol = file.FileName;
            textBox12.Text = tamYol;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (textBox10.Text != "")
            {
                string sorgu2 = "UPDATE Renkler SET renkyolu=@p1 WHERE RenkAdı=@RenkAdı";
                SqlCommand cmd2;
                cmd2 = new SqlCommand(sorgu2, bgl.baglanti());
                cmd2.Parameters.AddWithValue("@RenkAdı", textBox10.Text);
                cmd2.Parameters.AddWithValue("@p1", textBox12.Text);
                cmd2.ExecuteNonQuery();

                string sorgu3 = "UPDATE Stoklar SET renkyolu=@p1 WHERE Malzeme=@Malzeme";
                SqlCommand cmd3;
                cmd3 = new SqlCommand(sorgu3, bgl.baglanti());
                cmd3.Parameters.AddWithValue("@Malzeme", textBox10.Text);
                cmd3.Parameters.AddWithValue("@p1", textBox12.Text);
                cmd3.ExecuteNonQuery();
                MessageBox.Show("Renk başarıyla güncellendi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);

                liste();
                liste2();
                malzemebilgigetir();
            }
            else
            {
                MessageBox.Show("Renk güncelleyebilmek için tüm alanları doldurunuz!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }
    }
}
