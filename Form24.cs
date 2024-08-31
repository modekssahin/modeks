using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Modeks
{
    public partial class Form24 : Form
    {
        public Form24()
        {
            InitializeComponent();
        }
        sqlsinif bgl = new sqlsinif();
        public string kullaniciadi;
        public string yetki;
        private void KayitSayisi()
        {
            label23.Text = "Listedeki Kayıt Sayısı : " + Convert.ToString(dataGridView1.Rows.Count - 1);
        }
        private async Task listeAsync()
        {
            string kayit = "SELECT * FROM SIPARISLER";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti_eski());
            komut.Parameters.AddWithValue("@p1", "Evet");
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            await Task.Run(() =>
            {
                da.Fill(dt);
            });
            dataGridView1.DataSource = dt;
            KayitSayisi();
        }
        private async Task SiparisNoyaGoreSiralaAsync()
        {
            string sipno = textBox1.Text;
            string sorgu = "SELECT * From SIPARISLER WHERE SIPARISNO Like '" + sipno + "%'";
            DataSet ds = new DataSet();

            await Task.Run(() =>
            {
                using (SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti_eski()))
                {
                    adap.Fill(ds, "Siparişler");
                }
            });

            dataGridView1.DataSource = ds.Tables[0];
            KayitSayisi();
        }
        private async Task MusteriAdinaGoreSiralaAsync()
        {
            string mus = textBox2.Text;
            string sorgu = "SELECT * From SIPARISLER WHERE MUSTERI Like '" + mus + "%' AND RENK Like '" + textBox3.Text + "%'";
            DataSet ds = new DataSet();

            await Task.Run(() =>
            {
                using (SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti_eski()))
                {
                    adap.Fill(ds, "Siparişler");
                }
            });

            dataGridView1.DataSource = ds.Tables[0];
            KayitSayisi();
        }
        private async Task M2GoreSiralaAsync()
        {
            string mus = textBox5.Text;
            string sorgu = "SELECT * From SIPARISLER WHERE M2 Like '" + mus + "%' AND RENK Like '" + textBox3.Text + "%'";
            DataSet ds = new DataSet();

            await Task.Run(() =>
            {
                using (SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti_eski()))
                {
                    adap.Fill(ds, "Siparişler");
                }
            });

            dataGridView1.DataSource = ds.Tables[0];
            KayitSayisi();
        }
        private async Task KapakAdetGoreSiralaAsync()
        {
            string mus = textBox4.Text;
            string sorgu = "SELECT * From SIPARISLER WHERE ADET Like '" + mus + "%' AND RENK Like '" + textBox3.Text + "%'";
            DataSet ds = new DataSet();

            await Task.Run(() =>
            {
                using (SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti_eski()))
                {
                    adap.Fill(ds, "Siparişler");
                }
            });

            dataGridView1.DataSource = ds.Tables[0];
            KayitSayisi();
        }
        private async Task AciklamaGoreSiralaAsync()
        {
            string mus = textBox6.Text;
            string sorgu = "SELECT * From SIPARISLER WHERE ACIKLAMA Like'" + mus + "%' AND RENK Like '" + textBox3.Text + "%'";
            DataSet ds = new DataSet();

            await Task.Run(() =>
            {
                using (SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti_eski()))
                {
                    adap.Fill(ds, "Siparişler");
                }
            });

            dataGridView1.DataSource = ds.Tables[0];
            KayitSayisi();
        }
        private async Task OnayliSiparisleriGetir()
        {
            string sorgu = "SELECT * From SIPARISLER WHERE ONAY is not null";
            DataSet ds = new DataSet();

            await Task.Run(() =>
            {
                using (SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti_eski()))
                {
                    adap.Fill(ds, "Siparişler");
                }
            });

            dataGridView1.DataSource = ds.Tables[0];
            KayitSayisi();
        }
        private async Task OnaysizSiparisleriGetir()
        {
            string sorgu = "SELECT * From SIPARISLER WHERE ONAY is null";
            DataSet ds = new DataSet();

            await Task.Run(() =>
            {
                using (SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti_eski()))
                {
                    adap.Fill(ds, "Siparişler");
                }
            });

            dataGridView1.DataSource = ds.Tables[0];
            KayitSayisi();
        }
        private async Task RengeGoreSiralaAsync()
        {
            string renk = textBox3.Text;
            string sorgu = "SELECT * From SIPARISLER WHERE RENK Like '" + renk + "%' AND MUSTERI Like '" + textBox2.Text + "%'";
            DataSet ds = new DataSet();

            await Task.Run(() =>
            {
                using (SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti_eski()))
                {
                    adap.Fill(ds, "Siparişler");
                }
            });

            dataGridView1.DataSource = ds.Tables[0];
            KayitSayisi();
        }
        private async void Form24_Load(object sender, EventArgs e)
        {
            await listeAsync();
        }

        private async void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
                await listeAsync();
            else
                await SiparisNoyaGoreSiralaAsync();
        }
        private async void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (textBox2.Text == "" && textBox3.Text == "")
                await listeAsync();
            else
                await MusteriAdinaGoreSiralaAsync();
        }

        private async void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (textBox3.Text == "" && textBox2.Text == "")
                await listeAsync();
            else
                await RengeGoreSiralaAsync();
        }

        private async void button18_Click(object sender, EventArgs e)
        {
            await OnayliSiparisleriGetir();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            await OnaysizSiparisleriGetir();
        }
        string siparişno,musteri,siparistarihi;
        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            siparişno = dataGridView1.CurrentRow.Cells["SIPARISNO"].Value.ToString();
            musteri = dataGridView1.CurrentRow.Cells["MUSTERI"].Value.ToString();
            siparistarihi = dataGridView1.CurrentRow.Cells["SIPARISTARIHI"].Value.ToString();
            label4.Text = "Seçilen Sipariş No: " + siparişno;
            label5.Text = "Seçilen Müşteri : " + musteri;
            label6.Text = "Sipariş Tarihi : " + siparistarihi;
        }

        private async void textBox5_TextChanged(object sender, EventArgs e)
        {
            if (textBox5.Text == "")
                await listeAsync();
            else
                await M2GoreSiralaAsync();
        }

        private async void textBox4_TextChanged(object sender, EventArgs e)
        {
            if (textBox4.Text == "")
                await listeAsync();
            else
                await KapakAdetGoreSiralaAsync();
        }

        private async void textBox6_TextChanged(object sender, EventArgs e)
        {
            if (textBox6.Text == "")
                await listeAsync();
            else
                await AciklamaGoreSiralaAsync();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form17 frm = new Form17();
            frm.kullaniciadi = kullaniciadi;
            frm.yetki = yetki;
            this.Hide();
            frm.ShowDialog();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Form10 frm = new Form10();
            frm.hangiformdanModeksEski = "true";
            frm.SiparisNoSuEski = siparişno;
            frm.MusteriEski = musteri;
            frm.SiparisTarihi = siparistarihi;
            frm.kullaniciadi = kullaniciadi;
            frm.yetki = yetki;
            this.Hide();
            frm.ShowDialog();
        }
    }
}
