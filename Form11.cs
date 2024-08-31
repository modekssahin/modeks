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
    public partial class Form11 : Form
    {
        public Form11()
        {
            InitializeComponent();
        }
        sqlsinif bgl = new sqlsinif();
        private void Form11_Load(object sender, EventArgs e)
        {
            this.ActiveControl = textBox2;
        }

        private void Form11_Activated(object sender, EventArgs e)
        {

        }
        private void PaletTasi()
        {
            Form7 frm = (Form7)Application.OpenForms["Form7"];
            frm.liste();
            frm.aynısipnugetirme();
            if (textBox2.Text != "" && textBox1.Text != "")
            {
                for (int i = 0; i < frm.dataGridView1.Rows.Count - 1; i++)
                {
                    if (textBox2.Text == frm.dataGridView1.Rows[i].Cells[9].Value.ToString())
                    {
                        izin = "var";
                        sipno = frm.dataGridView1.Rows[i].Cells[0].Value.ToString();
                        renk = frm.dataGridView1.Rows[i].Cells[3].Value.ToString();
                        if (izin == "var")
                        {
                            string sorgu = "UPDATE Paletler SET Palet=@Palet WHERE Renk=@Renk ";
                            SqlCommand komut;
                            komut = new SqlCommand(sorgu, bgl.baglanti());
                            komut.Parameters.AddWithValue("@Renk", renk);
                            komut.Parameters.AddWithValue("@Palet", textBox1.Text);
                            komut.ExecuteNonQuery();

                            Form7 frm2 = (Form7)Application.OpenForms["Form7"];
                            frm2.liste();
                            frm2.aynısipnugetirme();

                            string sorgu2 = "UPDATE Siparişler SET Palet=@Palet WHERE SiparisNo=@SiparisNo AND KesildiMi='" + "Evet" + "' AND AnaSiparişMi='" + "Evet" + "' ";
                            SqlCommand komut2;
                            komut2 = new SqlCommand(sorgu2, bgl.baglanti());
                            komut2.Parameters.AddWithValue("@SiparisNo", sipno);
                            komut2.Parameters.AddWithValue("@Palet", textBox1.Text);
                            komut2.ExecuteNonQuery();
                            başarılı = "1";
                            Form7 frm3 = (Form7)Application.OpenForms["Form7"];
                            frm3.liste();
                            frm3.aynısipnugetirme();
                        }
                    }
                    izin = "";
                }
            }
            else
            {
                MessageBox.Show("Lütfen alanları doldurunuz!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            if (başarılı == "1")
            {
                Form7 frm2 = (Form7)Application.OpenForms["Form7"];
                MessageBox.Show("Palet Başarıyla Taşınmıştır!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                frm2.liste();
                frm2.aynısipnugetirme();
            }
        }
        string izin;
        string sipno;
        string renk;
        string başarılı;
        public void button9_Click(object sender, EventArgs e)
        {
            PaletTasi();
        }

        private void Form11_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        private void Form11_FormClosed(object sender, FormClosedEventArgs e)
        {
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                PaletTasi();
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                PaletTasi();
            }
        }
    }
}
