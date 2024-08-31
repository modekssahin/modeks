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
    public partial class Form14 : Form
    {
        public Form14()
        {
            InitializeComponent();
        }
        sqlsinif bgl = new sqlsinif();
        public void liste()
        {
            string kayit = "SELECT DISTINCT SiparisNo,Müşteri,SiparişTarihi,OnayTarihi,Boy,En,ToplamAdet,ToplamM2,Palet From Siparişler where Onay=@Onay ORDER BY SiparisNo ASC";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            komut.Parameters.AddWithValue("@Onay", "Onaylandı");
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
        }
        private void Form14_Load(object sender, EventArgs e)
        {
            liste();
            aynısipnugetirme();
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
            }
            catch (Exception)
            {

            }
        }
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            string srg = textBox3.Text;
            string sorgu = "SELECT DISTINCT SiparisNo,Müşteri,SiparişTarihi,Boy,En,ToplamAdet,ToplamM2,Palet From Siparişler where Boy Like '" + srg + "' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            string srg = textBox4.Text;
            string sorgu = "SELECT DISTINCT SiparisNo,Müşteri,SiparişTarihi,Boy,En,ToplamAdet,ToplamM2,Palet From Siparişler where En Like '" + srg + "' ORDER BY SiparisNo ASC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Siparişler");
            this.dataGridView1.DataSource = ds.Tables[0];
        }

        private void button12_Click(object sender, EventArgs e)
        {
            liste();
            aynısipnugetirme();
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            if (textBox3.Text != "")
            {
                string srg = (Convert.ToDouble(textBox3.Text) + Convert.ToDouble(numericUpDown1.Value)).ToString();
                string sorgu = "SELECT DISTINCT SiparisNo,Müşteri,SiparişTarihi,Boy,En,ToplamAdet,ToplamM2,Palet From Siparişler where Boy Like '" + srg + "' ORDER BY SiparisNo ASC";
                SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
                DataSet ds = new DataSet();
                adap.Fill(ds, "Siparişler");
                this.dataGridView1.DataSource = ds.Tables[0];
            }
            if (textBox4.Text != "")
            {
                string srg2 = (Convert.ToDouble(textBox4.Text) + Convert.ToDouble(numericUpDown1.Value)).ToString();
                string sorgu2 = "SELECT DISTINCT SiparisNo,Müşteri,SiparişTarihi,Boy,En,ToplamAdet,ToplamM2,Palet From Siparişler where En Like '" + srg2 + "' ORDER BY SiparisNo ASC";
                SqlDataAdapter adap2 = new SqlDataAdapter(sorgu2, bgl.baglanti());
                DataSet ds2 = new DataSet();
                adap2.Fill(ds2, "Siparişler");
                this.dataGridView1.DataSource = ds2.Tables[0];
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTime bitir = dateTimePicker2.Value;
            DateTime basla = dateTimePicker1.Value;
            label4.Text = basla.ToString("yyyy - MM - dd");
            label7.Text = bitir.ToString("yyyy - MM - dd  HH:mm:ss");
            string kayit = "SELECT DISTINCT SiparisNo,Müşteri,SiparişTarihi,OnayTarihi,Boy,En,ToplamAdet,ToplamM2,Palet From Siparişler where Onay=@Onay AND OnayTarihi between '" + label4.Text + "' AND '" + label7.Text + "' ORDER BY SiparisNo ASC";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            komut.Parameters.AddWithValue("@Onay", "Onaylandı");
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            aynısipnugetirme();
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            DateTime bitir = dateTimePicker2.Value;
            DateTime basla = dateTimePicker1.Value;
            label4.Text = basla.ToString("yyyy - MM - dd");
            label7.Text = bitir.ToString("yyyy - MM - dd  HH:mm:ss");
            string kayit = "SELECT DISTINCT SiparisNo,Müşteri,SiparişTarihi,OnayTarihi,Boy,En,ToplamAdet,ToplamM2,Palet From Siparişler where Onay=@Onay AND OnayTarihi between '" + label4.Text + "' AND '" + label7.Text + "' ORDER BY SiparisNo ASC";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            komut.Parameters.AddWithValue("@Onay", "Onaylandı");
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            aynısipnugetirme();

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
