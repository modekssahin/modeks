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
using System.Windows.Forms.DataVisualization.Charting;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Modeks
{
    public partial class Form13 : Form
    {
        public Form13()
        {
            InitializeComponent();
        }
        sqlsinif bgl = new sqlsinif();
        double toplamm2;
        double toplamkesilen;
        double toplamkesilecek;
        //int x = 0;
        public void liste()
        {
            string kayit = "SELECT * FROM RenkDurum ORDER BY PressHazirM2 DESC";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
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
        private void matrenkler()
        {
            string srg = "HG%";
            string sorgu = "SELECT * FROM RenkDurum WHERE Renk not like 'HG%' ORDER BY PressHazirM2 DESC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Renkler");
            this.dataGridView1.DataSource = ds.Tables[0];
            aynısipnugetirme();
            Hesapla();
        }
        private void parlakrenkler()
        {
            string srg = "HG%";
            string sorgu = "SELECT * FROM RenkDurum WHERE Renk like 'HG%' ORDER BY PressHazirM2 DESC";
            SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
            DataSet ds = new DataSet();
            adap.Fill(ds, "Renkler");
            this.dataGridView1.DataSource = ds.Tables[0];
            aynısipnugetirme();
            Hesapla();
        }
        private void Hesapla()
        {
            toplamm2 = 0;
            toplamkesilen = 0;
            toplamkesilecek = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                toplamm2 += Convert.ToDouble(dataGridView1.Rows[i].Cells["ToplamM2"].Value.ToString());
                toplamkesilen += Convert.ToDouble(dataGridView1.Rows[i].Cells["PressHazirM2"].Value.ToString());
                toplamkesilecek += Convert.ToDouble(dataGridView1.Rows[i].Cells["PaletlenecekM2"].Value.ToString());

            }
            button1.Text = toplamm2.ToString();
            button3.Text = toplamkesilen.ToString();
            button4.Text = toplamkesilecek.ToString();
        }
        private void Form13_Load(object sender, EventArgs e)
        {
            liste();
            aynısipnugetirme();
            Hesapla();
            renklendir();
            button2.Text = dataGridView1.Rows.Count - 1 + " adet renk";
        }

        private void button6_Click(object sender, EventArgs e)
        {
            matrenkler();
            renklendir();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            parlakrenkler();
            renklendir();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            liste();
            aynısipnugetirme();
            Hesapla();
            renklendir();
        }

        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            aynısipnugetirme();
            renklendir();
        }
        private void renklendir()
        {
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (Convert.ToDouble(dataGridView1.Rows[i].Cells["PressHazirM2"].Value.ToString()) > 50)
                {
                    dataGridView1.Rows[i].Cells["PressHazirM2"].Style.BackColor = Color.SeaGreen;
                    dataGridView1.Rows[i].Cells["PressHazirM2"].Style.ForeColor = Color.White;
                }
            }
        }
        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void Form13_Shown(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {

        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            Grafik p = new Grafik();
            p.Show();
        }
    }
}

