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
    public partial class Form15 : Form
    {
        public Form15()
        {
            InitializeComponent();
        }
        sqlsinif bgl = new sqlsinif();
        public string müşteri;
        public string renk;
        public void liste()
        {
            string kayit = "SELECT SiparisNo,Müşteri,Model,Renk,SiparişTarihi,ToplamM2 From Siparişler WHERE Müşteri=@Müşteri AND Renk=@Renk ORDER BY SiparisNo ASC";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            komut.Parameters.AddWithValue("@Müşteri", müşteri);
            komut.Parameters.AddWithValue("@Renk", renk);
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
        }
        public void tümkayıtlar()
        {
            string kayit = "SELECT SiparisNo,Müşteri,Model,Renk,SiparişTarihi,ToplamM2 From Siparişler WHERE Müşteri=@Müşteri ORDER BY SiparisNo ASC";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            komut.Parameters.AddWithValue("@Müşteri", müşteri);
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
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
        private void Form15_Load(object sender, EventArgs e)
        {
            liste();
            aynısipnugetirme();
            textBox3.Text = müşteri;
            textBox1.Text = renk;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            tümkayıtlar();
            aynısipnugetirme();
        }

        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            aynısipnugetirme();
        }
    }
}
