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
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
        }
        sqlsinif bgl = new sqlsinif();
        public string yetki;
        string id;
        private void liste()
        {
            string kayit = "SELECT * from Modeller";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
        }
        private void cmbox1doldur()
        {
            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT *FROM M2_Kapak_Farkı";
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                comboBox1.Items.Add(dr["Değer"]);
            }
        }
        private void cmbox2doldur()
        {
            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT *FROM Acil_Fiyatı";
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                comboBox2.Items.Add(dr["Değer"]);
            }
        }
        private void cmbox3doldur()
        {
            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT *FROM M2_Kapak_Farkı_2";
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                comboBox3.Items.Add(dr["Değer"]);
            }
        }
        private void modelegöresırala()
        {

            string srg = textBox1.Text;
            string sorgu = "SELECT * from Modeller where ModelAdı Like '%" + srg + "%'";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Modeller");
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
        private void Form4_Load(object sender, EventArgs e)
        {
            timer1.Start();
            liste();
            cmbox1doldur();
            cmbox2doldur();
            cmbox3doldur();
            SqlCommand komut = new SqlCommand();
            komut.CommandText = "SELECT *FROM Kargo_Fiyatı where id=@p1";
            komut.Parameters.AddWithValue("@p1", 1);
            komut.Connection = bgl.baglanti();
            komut.CommandType = CommandType.Text;

            SqlDataReader dr;
            dr = komut.ExecuteReader();
            while (dr.Read())
            {
                textBox10.Text = Convert.ToDouble(dr["KargoFiyatı"]).ToString();
                textBox33.Text = Convert.ToDouble(dr["KargoFiyatiKucukse"]).ToString();
                textBox34.Text = Convert.ToDouble(dr["KargoFiyatiKucukKucukse"]).ToString();

            }
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            string kayit = "insert into Modeller(ModelAdı,HgFiyat,MaTFiyat,Cam,[Cam 2 Kafes],[Cam 4 Kafes],[Cam 6 Kafes],[Cam 8 Kafes],resimyolu,RSM,SOFT,ULTRASOFT,BASILMICAK)values (@p1,@p2,@p3,@p4,@p5,@p6,@p7,@p8,@p9,@p10,@p11,@p12,@p13)";
            SqlCommand cmd = new SqlCommand(kayit, bgl.baglanti());
            cmd.Parameters.AddWithValue("@p1", textBox3.Text);
            cmd.Parameters.AddWithValue("@p2", textBox4.Text);
            cmd.Parameters.AddWithValue("@p3", textBox5.Text);
            cmd.Parameters.AddWithValue("@p4", textBox6.Text);
            cmd.Parameters.AddWithValue("@p5", textBox2.Text);
            cmd.Parameters.AddWithValue("@p6", textBox7.Text);
            cmd.Parameters.AddWithValue("@p7", textBox8.Text);
            cmd.Parameters.AddWithValue("@p8", textBox9.Text);
            cmd.Parameters.AddWithValue("@p9", textBox22.Text);
            cmd.Parameters.AddWithValue("@p10", textBox30.Text);
            cmd.Parameters.AddWithValue("@p11", textBox29.Text);
            cmd.Parameters.AddWithValue("@p12", textBox28.Text);
            cmd.Parameters.AddWithValue("@p13", textBox31.Text);
            cmd.ExecuteNonQuery();
            MessageBox.Show("Model Ekleme İşlemi Gerçekleşti.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            liste();
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            id = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            textBox15.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            textBox20.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            textBox21.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            textBox19.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            textBox18.Text = dataGridView1.CurrentRow.Cells["Cam"].Value.ToString();
            textBox17.Text = dataGridView1.CurrentRow.Cells["Cam 2 Kafes"].Value.ToString();
            textBox12.Text = dataGridView1.CurrentRow.Cells["Cam 4 Kafes"].Value.ToString();
            textBox13.Text = dataGridView1.CurrentRow.Cells["Cam 6 Kafes"].Value.ToString();
            textBox14.Text = dataGridView1.CurrentRow.Cells["Cam 8 Kafes"].Value.ToString();
            textBox26.Text = dataGridView1.CurrentRow.Cells["RSM"].Value.ToString();
            textBox25.Text = dataGridView1.CurrentRow.Cells["SOFT"].Value.ToString();
            textBox27.Text = dataGridView1.CurrentRow.Cells["ULTRASOFT"].Value.ToString();
            textBox32.Text = dataGridView1.CurrentRow.Cells["BASILMICAK"].Value.ToString();
            textBox23.Text = dataGridView1.CurrentRow.Cells["resimyolu"].Value.ToString();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string sorgu = "UPDATE Modeller SET ModelAdı=@p1,HgFiyat=@p2,MaTFiyat=@p3,Cam=@p4,[Cam 2 Kafes]=@p5,[Cam 4 Kafes]=@p6,[Cam 6 Kafes]=@p7,[Cam 8 Kafes]=@p8, resimyolu=@p9, RSM=@p10, SOFT=@p11, ULTRASOFT=@p12, BASILMICAK=@p13 WHERE id=@id";
            SqlCommand cmd;
            cmd = new SqlCommand(sorgu, bgl.baglanti());
            cmd.Parameters.AddWithValue("@id", Convert.ToInt32(id.ToString()));
            cmd.Parameters.AddWithValue("@p1", textBox20.Text);
            cmd.Parameters.AddWithValue("@p2", textBox21.Text);
            cmd.Parameters.AddWithValue("@p3", textBox19.Text);
            cmd.Parameters.AddWithValue("@p4", textBox18.Text);
            cmd.Parameters.AddWithValue("@p5", textBox17.Text);
            cmd.Parameters.AddWithValue("@p6", textBox12.Text);
            cmd.Parameters.AddWithValue("@p7", textBox13.Text);
            cmd.Parameters.AddWithValue("@p8", textBox14.Text);
            cmd.Parameters.AddWithValue("@p9", textBox23.Text);
            cmd.Parameters.AddWithValue("@p10", textBox26.Text);
            cmd.Parameters.AddWithValue("@p11", textBox25.Text);
            cmd.Parameters.AddWithValue("@p12", textBox27.Text);
            cmd.Parameters.AddWithValue("@p13", textBox32.Text);
            cmd.ExecuteNonQuery();
            MessageBox.Show("Model Güncelleme İşlemi Gerçekleşti.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            liste();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                liste();
            }
            modelegöresırala();
        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(comboBox1.Text!="" && textBox24.Text != "")
            {
                string sorgu = "UPDATE M2_Kapak_Farkı SET Yüzde=@p1 WHERE Değer=@Değer";
                SqlCommand cmd;
                cmd = new SqlCommand(sorgu, bgl.baglanti());
                cmd.Parameters.AddWithValue("@Değer", comboBox1.Text);
                cmd.Parameters.AddWithValue("@p1", textBox24.Text);
                cmd.ExecuteNonQuery();
                MessageBox.Show("M2 Kapak Farkı Güncelleme İşlemi Gerçekleşti.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Lütfen listeden bir değer seçiniz ve yüzde giriniz. Yüzde kısmına Örneğin 40 olarak giriş yapınız.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox10.Text != "" && textBox33.Text != "")
            {

                string sorgu = "UPDATE Kargo_Fiyatı SET KargoFiyatı=@p1, KargoFiyatiKucukse=@p2, KargoFiyatiKucukKucukse=@p3  WHERE id=@id";
                SqlCommand cmd;
                cmd = new SqlCommand(sorgu, bgl.baglanti());
                cmd.Parameters.AddWithValue("@id", 1);
                cmd.Parameters.AddWithValue("@p1", textBox10.Text);
                cmd.Parameters.AddWithValue("@p2", textBox33.Text);
                cmd.Parameters.AddWithValue("@p3", textBox34.Text);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Kargo Fiyatı Güncelleme İşlemi Gerçekleşti.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Lütfen bir kargo fiyatlarını giriniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox2.Text != "" && textBox11.Text != "")
            {
                string sorgu = "UPDATE Acil_Fiyatı SET Yüzde=@p1 WHERE Değer=@Değer";
                SqlCommand cmd;
                cmd = new SqlCommand(sorgu, bgl.baglanti());
                cmd.Parameters.AddWithValue("@Değer", comboBox2.Text);
                cmd.Parameters.AddWithValue("@p1", textBox11.Text);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Acil Fiyatı Güncelleme İşlemi Gerçekleşti.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Lütfen listeden bir değer seçiniz ve yüzde giriniz. Yüzde kısmına Örneğin 40 olarak giriş yapınız.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (textBox15.Text!="")
            {
                DialogResult d1 = new DialogResult();
                d1 = MessageBox.Show("Silmek istediğinizden emin misiniz? ", "Uyarı", MessageBoxButtons.YesNo);
                if (d1 == DialogResult.Yes)
                {
                    string sorgu = "DELETE FROM Modeller WHERE id=@id";
                    SqlCommand komut;
                    komut = new SqlCommand(sorgu, bgl.baglanti());
                    komut.Parameters.AddWithValue("@id", Convert.ToInt32(textBox15.Text));
                    komut.ExecuteNonQuery();
                    MessageBox.Show("Model başarıyla silindi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    liste();
                }
            }
            else
            {
                MessageBox.Show("Model Silebilmek İçin listeden bir kayıt seçiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void Form4_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (comboBox3.Text != "" && textBox16.Text != "")
            {
                string sorgu = "UPDATE M2_Kapak_Farkı_2 SET Yüzde=@p1 WHERE Değer=@Değer";
                SqlCommand cmd;
                cmd = new SqlCommand(sorgu, bgl.baglanti());
                cmd.Parameters.AddWithValue("@Değer", comboBox3.Text);
                cmd.Parameters.AddWithValue("@p1", textBox16.Text);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Güncelleme İşlemi Gerçekleşti.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Lütfen listeden bir değer seçiniz ve yüzde giriniz. Yüzde kısmına Örneğin 40 olarak giriş yapınız.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "Resim Formatı |*.png; *.jpg; *.pjeg" ;
            file.ShowDialog();
            string tamYol = file.FileName;
            textBox22.Text = tamYol;
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void button9_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "Resim Formatı |*.png; *.jpg; *.pjeg";
            file.ShowDialog();
            string tamYol = file.FileName;
            textBox23.Text = tamYol;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            textBox23.Text = "";
        }

        private void button12_Click(object sender, EventArgs e)
        {

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
