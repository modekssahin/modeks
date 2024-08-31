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
    public partial class Form16 : Form
    {
        public Form16()
        {
            InitializeComponent();
        }
        public string yetki;
        sqlsinif bgl = new sqlsinif();
        public void liste()
        {
            string kayit = "SELECT * FROM RenkDurum";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
        }
        private void Form16_Load(object sender, EventArgs e)
        {
            liste();
            button2.Text = dataGridView1.Rows.Count - 1 + " adet renk";
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
