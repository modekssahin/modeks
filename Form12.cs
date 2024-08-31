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
using System.Threading;
using System.Windows.Controls.Primitives;

namespace Modeks
{
    public partial class Form12 : Form
    {
        public Form12()
        {
            InitializeComponent();
        }
        sqlsinif bgl = new sqlsinif();
        public void liste()
        {
            string kayit = "SELECT DISTINCT KesildiMi,SiparisNo,Müşteri,Model,TeslimTarihi,KesildiTarihi,Renk,Palet,ToplamM2 From Siparişler where /*KesildiMi=@KesildiMi AND */AnaSiparişMi=@p1 AND Palet=@Palet AND Aşama=@Aşama ORDER BY SiparisNo ASC";
            SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
            //komut.Parameters.AddWithValue("@KesildiMi", "Evet");
            komut.Parameters.AddWithValue("@p1", "Evet");
            komut.Parameters.AddWithValue("@Palet", textBox1.Text);
            komut.Parameters.AddWithValue("@Aşama", "Etiket");
            SqlDataAdapter da = new SqlDataAdapter(komut);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
        }
        public void liste2()
        {
            if (dataGridView1.Rows.Count > 1)
            {
                //DateTime bitir = DateTime.Now;
                //DateTime basla = bitir.AddDays(-5);
                //label3.Text = basla.ToString("yyyy - MM - dd HH:mm:ss");
                //label4.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");
                //string kayit = "SELECT DISTINCT Onay,KesildiMi,SiparisNo,Müşteri,Model,TeslimTarihi,KesildiTarihi,Renk,Palet,Aşama,SiparişTarihi,ToplamM2 From Siparişler where Palet is null AND SiparişTarihi between '" + label3.Text + "' AND '" + label4.Text + "' AND AnaSiparişMi='Evet' AND Renk=@Renk ORDER BY SiparisNo ASC";
                //SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
                //komut.Parameters.AddWithValue("@p1", "Evet");
                //komut.Parameters.AddWithValue("@Renk", dataGridView1.Rows[0].Cells["Renk"].Value.ToString());
                //SqlDataAdapter da2 = new SqlDataAdapter(komut);
                //DataTable dt2 = new DataTable();
                //da2.Fill(dt2);
                //dataGridView2.DataSource = dt2;



                DateTime bitir = DateTime.Now;
                DateTime basla = bitir.AddDays(-5);
                label3.Text = basla.ToString("yyyy - MM - dd HH:mm:ss");
                label4.Text = bitir.ToString("yyyy - MM - dd HH:mm:ss");

                string kayit = "SELECT DISTINCT Onay,KesildiMi,SiparisNo,Müşteri,Model,TeslimTarihi,KesildiTarihi,Renk,Palet,Aşama,SiparişTarihi,ToplamM2 From Siparişler where AnaSiparişMi='Evet' AND Renk=@Renk AND Aşama in ('Onay Bekliyor', 'Onaylandı', 'Etiket', 'Palet') AND SiparisNo != @SiparisNo ORDER BY SiparisNo ASC";
                SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
                komut.Parameters.AddWithValue("@p1", "Evet");
                komut.Parameters.AddWithValue("@Renk", dataGridView1.Rows[0].Cells["Renk"].Value.ToString());
                komut.Parameters.AddWithValue("@SiparisNo", dataGridView1.Rows[0].Cells["SiparisNo"].Value.ToString());
                SqlDataAdapter da2 = new SqlDataAdapter(komut);
                DataTable dt2 = new DataTable();
                da2.Fill(dt2);
                dataGridView2.DataSource = dt2;
            }
            else
            {
                string kayit = "SELECT DISTINCT Onay,KesildiMi,SiparisNo,Müşteri,Model,TeslimTarihi,KesildiTarihi,Renk,Palet,Aşama,SiparişTarihi,ToplamM2 From Siparişler where Palet is null AND AnaSiparişMi=@p1 AND Palet=@Palet ORDER BY SiparisNo ASC";

                SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
                komut.Parameters.AddWithValue("@p1", "Evet");
                komut.Parameters.AddWithValue("@Palet", textBox1.Text);
                SqlDataAdapter da2 = new SqlDataAdapter(komut);
                DataTable dt2 = new DataTable();
                da2.Fill(dt2);
                dataGridView2.DataSource = dt2;
            }

        }
        private void aynısipnugetirme()
        {
            try
            {
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < i; j++)
                    {
                        if (dataGridView1.Rows[j].Cells[1].Value.ToString() == dataGridView1.Rows[i].Cells[1].Value.ToString())
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
        private void aynısipnugetirme2()
        {
            try
            {
                for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < i; j++)
                    {
                        if (dataGridView2.Rows[j].Cells["SiparisNo"].Value.ToString() == dataGridView2.Rows[i].Cells["SiparisNo"].Value.ToString())
                        {
                            dataGridView2.Rows[i].Visible = false;
                        }

                    }

                }
            }
            catch (Exception)
            {

            }
        }
        private void Form12_Load(object sender, EventArgs e)
        {
            checkBox1.Checked = true;
            checkBox2.Checked = false;
            liste();
            liste2();
            aynısipnugetirme();
            aynısipnugetirme2();
        }
        //private void üçgüniçindeki()
        //{
        //    DateTime bitir = DateTime.Now;
        //    DateTime basla = bitir.AddDays(-3);
        //    label3.Text = basla.ToString("yyyy - MM - dd");
        //    label4.Text = bitir.ToString("yyyy - MM - dd");
        //    string sorgu = "SELECT DISTINCT Onay,KesildiMi,SiparisNo,Müşteri,Model,TeslimTarihi,KesildiTarihi,Renk,Palet,Aşama,SiparişTarihi From Siparişler where Palet is null AND SiparişTarihi between '" + label3.Text + "' AND '" + label4.Text + "' AND AnaSiparişMi='Evet' ORDER BY SiparisNo ASC";
        //    SqlDataAdapter adap = new SqlDataAdapter(sorgu, bgl.baglanti());
        //    DataSet ds = new DataSet();
        //    adap.Fill(ds, "Siparişler");
        //    this.dataGridView2.DataSource = ds.Tables[0];
        //    aynısipnugetirme2();
        //}
        string hepsi;
        string paletlenen;
        string kesilecek;

        private void PaletiPresseGonder()
        {
            hepsi = "";
            paletlenen = "";
            kesilecek = "";
            if (dataGridView1.Rows.Count > 1)
            {
                if (dataGridView2.Rows.Count > 1)
                {
                    DialogResult dialog = new DialogResult();
                    dialog = MessageBox.Show("Kesilmeyen siparişler var! Paleti press'e gönderme işlemi onaylansın mı?", "ÇIKIŞ", MessageBoxButtons.YesNo);
                    if (dialog == DialogResult.Yes)
                    {

                        string sorgu = "UPDATE Siparişler SET Aşama=@Aşama WHERE Palet=@Palet AND Onay=@Onay AND KesildiMi=@KesildiMi";
                        SqlCommand komut;
                        komut = new SqlCommand(sorgu, bgl.baglanti());
                        //komut.Parameters.AddWithValue("@SiparisNo", dataGridView1.Rows[0].Cells[1].Value.ToString());
                        komut.Parameters.AddWithValue("@Palet", dataGridView1.Rows[0].Cells[7].Value.ToString());
                        komut.Parameters.AddWithValue("@Onay", "Onaylandı");
                        komut.Parameters.AddWithValue("@Aşama", "Palet");
                        komut.Parameters.AddWithValue("@KesildiMi", "Evet");
                        komut.ExecuteNonQuery();
                        MessageBox.Show(dataGridView1.Rows[0].Cells[7].Value.ToString() + " 'lu palet presse gönderilmiştir.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        string sorgu2 = "DELETE FROM Paletler WHERE Palet=@Palet";
                        SqlCommand komut2;
                        komut2 = new SqlCommand(sorgu2, bgl.baglanti());
                        komut2.Parameters.AddWithValue("@Palet", dataGridView1.Rows[0].Cells[7].Value.ToString());
                        komut2.ExecuteNonQuery();

                        for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                        {
                            if (dataGridView1.Rows[i].Visible == true)
                            {
                                SqlCommand komut3 = new SqlCommand();
                                komut3.CommandText = "SELECT *From Grafik where Renk=@Renk";
                                komut3.Parameters.AddWithValue("@Renk", dataGridView1.Rows[i].Cells["Renk"].Value.ToString());
                                komut3.Connection = bgl.baglanti();
                                komut3.CommandType = CommandType.Text;
                                SqlDataReader dr;
                                dr = komut3.ExecuteReader();
                                while (dr.Read())
                                {
                                    hepsi = dr["Hepsi"].ToString();
                                    paletlenen = dr["Paletlenen"].ToString();
                                }

                                string sorgu4 = "UPDATE Grafik SET Hepsi=@Hepsi, Paletlenen=@Paletlenen, Palet=@Palet WHERE Renk=@Renk";
                                SqlCommand komut4;
                                komut4 = new SqlCommand(sorgu4, bgl.baglanti());
                                komut4.Parameters.AddWithValue("@Renk", dataGridView1.Rows[i].Cells["Renk"].Value.ToString());
                                komut4.Parameters.AddWithValue("@Hepsi", Convert.ToDouble(Convert.ToDouble(hepsi) - Convert.ToDouble(dataGridView1.Rows[i].Cells["ToplamM2"].Value.ToString())));
                                komut4.Parameters.AddWithValue("@Paletlenen", 0);
                                komut4.Parameters.AddWithValue("@Palet", "");
                                komut4.ExecuteNonQuery();

                                if (dataGridView1.Rows[0].Cells["Renk"].Value.ToString() == "BASILMICAK")
                                {
                                    string sorgubasilmicak = "UPDATE Siparişler SET Aşama=@Aşama, MembranPressTarihi=@MembranPressTarihi WHERE SiparisNo=@SiparisNo";
                                    SqlCommand komutbasilmicak;
                                    komutbasilmicak = new SqlCommand(sorgubasilmicak, bgl.baglanti());
                                    komutbasilmicak.Parameters.AddWithValue("@SiparisNo", dataGridView1.Rows[i].Cells["SiparisNo"].Value.ToString());
                                    komutbasilmicak.Parameters.AddWithValue("@MembranPressTarihi", DateTime.Now);
                                    komutbasilmicak.Parameters.AddWithValue("@Aşama", "Kargo");
                                    komutbasilmicak.ExecuteNonQuery();
                                }
                            }
                        }
                       
                        textBox1.Text = "";
                        Form7 frm2 = (Form7)Application.OpenForms["Form7"];
                        frm2.liste();
                        frm2.aynısipnugetirme();
                        frm2.AcilSipariş();
                    }
                    else
                    {
                        MessageBox.Show("Gönderme işlemi iptal edildi");
                    }
                }
                else
                {
                   
                    string sorgu = "UPDATE Siparişler SET Aşama=@Aşama WHERE Palet=@Palet";
                    SqlCommand komut;
                    komut = new SqlCommand(sorgu, bgl.baglanti());
                    //komut.Parameters.AddWithValue("@SiparisNo", dataGridView1.Rows[0].Cells[1].Value.ToString());
                    komut.Parameters.AddWithValue("@Palet", dataGridView1.Rows[0].Cells[7].Value.ToString());
                    komut.Parameters.AddWithValue("@Aşama", "Palet");
                    komut.ExecuteNonQuery();
                    MessageBox.Show(dataGridView1.Rows[0].Cells[7].Value.ToString() + " 'lu palet presse gönderilmiştir.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    string sorgu2 = "DELETE FROM Paletler WHERE Palet=@Palet";
                    SqlCommand komut2;
                    komut2 = new SqlCommand(sorgu2, bgl.baglanti());
                    komut2.Parameters.AddWithValue("@Palet", dataGridView1.Rows[0].Cells[7].Value.ToString());
                    komut2.ExecuteNonQuery();

                    for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                    {
                        if (dataGridView1.Rows[i].Visible == true)
                        {
                            SqlCommand komut3 = new SqlCommand();
                            komut3.CommandText = "SELECT *From Grafik where Renk=@Renk";
                            komut3.Parameters.AddWithValue("@Renk", dataGridView1.Rows[i].Cells["Renk"].Value.ToString());
                            komut3.Connection = bgl.baglanti();
                            komut3.CommandType = CommandType.Text;
                            SqlDataReader dr;
                            dr = komut3.ExecuteReader();
                            while (dr.Read())
                            {
                                hepsi = dr["Hepsi"].ToString();
                                paletlenen = dr["Paletlenen"].ToString();
                                kesilecek = dr["Paletlenecek"].ToString();
                            }

                            string sorgu4 = "UPDATE Grafik SET Hepsi=@Hepsi, Paletlenen=@Paletlenen, Palet=@Palet WHERE Renk=@Renk";
                            SqlCommand komut4;
                            komut4 = new SqlCommand(sorgu4, bgl.baglanti());
                            komut4.Parameters.AddWithValue("@Renk", dataGridView1.Rows[i].Cells["Renk"].Value.ToString());
                            komut4.Parameters.AddWithValue("@Hepsi", Convert.ToDouble(Convert.ToDouble(hepsi) - Convert.ToDouble(dataGridView1.Rows[i].Cells["ToplamM2"].Value.ToString())));
                            komut4.Parameters.AddWithValue("@Paletlenen", 0);
                            komut4.Parameters.AddWithValue("@Palet", "");
                            komut4.ExecuteNonQuery();

                            if (dataGridView1.Rows[0].Cells["Renk"].Value.ToString() == "BASILMICAK")
                            {
                                string sorgubasilmicak = "UPDATE Siparişler SET Aşama=@Aşama, MembranPressTarihi=@MembranPressTarihi WHERE SiparisNo=@SiparisNo";
                                SqlCommand komutbasilmicak;
                                komutbasilmicak = new SqlCommand(sorgubasilmicak, bgl.baglanti());
                                komutbasilmicak.Parameters.AddWithValue("@SiparisNo", dataGridView1.Rows[i].Cells["SiparisNo"].Value.ToString());
                                komutbasilmicak.Parameters.AddWithValue("@MembranPressTarihi", DateTime.Now);
                                komutbasilmicak.Parameters.AddWithValue("@Aşama", "Kargo");
                                komutbasilmicak.ExecuteNonQuery();
                            }
                        }

                    }

                    textBox1.Text = "";
                    Form7 frm3 = (Form7)Application.OpenForms["Form7"];
                    frm3.liste();
                    frm3.aynısipnugetirme();
                    frm3.AcilSipariş();
                }


            }
        }
        
        private void SiparisiPresseGonder()
        {
            hepsi = "";
            paletlenen = "";
            kesilecek = "";
           
                        string sorgu = "UPDATE Siparişler SET Aşama=@Aşama WHERE SiparisNo=@SiparisNo AND Onay=@Onay AND KesildiMi=@KesildiMi";
                        SqlCommand komut;
                        komut = new SqlCommand(sorgu, bgl.baglanti());
                        komut.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
                        komut.Parameters.AddWithValue("@Onay", "Onaylandı");
                        komut.Parameters.AddWithValue("@Aşama", "Palet");
                        komut.Parameters.AddWithValue("@KesildiMi", "Evet");
                        komut.ExecuteNonQuery();
                        MessageBox.Show(textBox1.Text + " 'lu sipariş presse gönderilmiştir.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                  
            string sorguPalet = @"
SELECT Palet 
FROM Siparişler 
WHERE SiparisNo = @SiparisNo 
";

            string paletNumarasi = null;
            using (SqlCommand komutPalet = new SqlCommand(sorguPalet, bgl.baglanti()))
            {
                komutPalet.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
                paletNumarasi = (string)komutPalet.ExecuteScalar();
            }

            string sorguCount = @"
SELECT COUNT(*)
FROM Siparişler
WHERE
Renk IN (
    SELECT Renk
    FROM Siparişler 
    WHERE Palet = @paletNumarasi
    AND (Aşama = 'Etiket')
	GROUP BY Renk
	)
";

            int count = 0;
            using (SqlCommand komutCount = new SqlCommand(sorguCount, bgl.baglanti()))
            {
                komutCount.Parameters.AddWithValue("@paletNumarasi", paletNumarasi);
                count = (int)komutCount.ExecuteScalar();
            }

            if (count == 0 && paletNumarasi != null)
            {
                string sorgu2 = "DELETE FROM Paletler WHERE Palet=@Palet";
                using (SqlCommand komut2 = new SqlCommand(sorgu2, bgl.baglanti()))
                {
                    komut2.Parameters.AddWithValue("@Palet", paletNumarasi);
                    komut2.ExecuteNonQuery();
                }
            }

            /*
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    if (dataGridView1.Rows[i].Visible == true)
                    {
                        SqlCommand komut3 = new SqlCommand();
                        komut3.CommandText = "SELECT *From Grafik where Renk=@Renk";
                        komut3.Parameters.AddWithValue("@Renk", dataGridView1.Rows[i].Cells["Renk"].Value.ToString());
                        komut3.Connection = bgl.baglanti();
                        komut3.CommandType = CommandType.Text;
                        SqlDataReader dr;
                        dr = komut3.ExecuteReader();
                        while (dr.Read())
                        {
                            hepsi = dr["Hepsi"].ToString();
                            paletlenen = dr["Paletlenen"].ToString();
                        }

                        string sorgu4 = "UPDATE Grafik SET Hepsi=@Hepsi, Paletlenen=@Paletlenen, Palet=@Palet WHERE Renk=@Renk";
                        SqlCommand komut4;
                        komut4 = new SqlCommand(sorgu4, bgl.baglanti());
                        komut4.Parameters.AddWithValue("@Renk", dataGridView1.Rows[i].Cells["Renk"].Value.ToString());
                        komut4.Parameters.AddWithValue("@Hepsi", Convert.ToDouble(Convert.ToDouble(hepsi) - Convert.ToDouble(dataGridView1.Rows[i].Cells["ToplamM2"].Value.ToString())));
                        komut4.Parameters.AddWithValue("@Paletlenen", 0);
                        komut4.Parameters.AddWithValue("@Palet", "");
                        komut4.ExecuteNonQuery();

                        if (dataGridView1.Rows[0].Cells["Renk"].Value.ToString() == "BASILMICAK")
                        {
                            string sorgubasilmicak = "UPDATE Siparişler SET Aşama=@Aşama, MembranPressTarihi=@MembranPressTarihi WHERE SiparisNo=@SiparisNo";
                            SqlCommand komutbasilmicak;
                            komutbasilmicak = new SqlCommand(sorgubasilmicak, bgl.baglanti());
                            komutbasilmicak.Parameters.AddWithValue("@SiparisNo", dataGridView1.Rows[i].Cells["SiparisNo"].Value.ToString());
                            komutbasilmicak.Parameters.AddWithValue("@MembranPressTarihi", DateTime.Now);
                            komutbasilmicak.Parameters.AddWithValue("@Aşama", "Kargo");
                            komutbasilmicak.ExecuteNonQuery();
                        }
                    }
                }
            */
                        Form7 frm2 = (Form7)Application.OpenForms["Form7"];
                        frm2.liste();
                        frm2.aynısipnugetirme();
                        frm2.AcilSipariş();
     }
        private void button2_Click(object sender, EventArgs e)
        {
            if(checkBox1.Checked)
            {
                PaletiPresseGonder();
            } else if(checkBox2.Checked)
            {
                SiparisiPresseGonder();
                MessageBox.Show("Siparişi gönderdim.");
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if(checkBox1.Checked)
            {
                liste();
                liste2();
                aynısipnugetirme();
                aynısipnugetirme2();
            } else if(checkBox2.Checked)
            {
                string kayit = "SELECT DISTINCT KesildiMi,SiparisNo,Müşteri,Model,TeslimTarihi,KesildiTarihi,Renk,Palet,ToplamM2 From Siparişler where /*KesildiMi=@KesildiMi AND */AnaSiparişMi=@p1 AND SiparisNo=@SiparisNo AND Aşama=@Aşama ORDER BY SiparisNo ASC";
                SqlCommand komut = new SqlCommand(kayit, bgl.baglanti());
                komut.Parameters.AddWithValue("@p1", "Evet");
                komut.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
                komut.Parameters.AddWithValue("@Aşama", "Etiket");
                SqlDataAdapter da = new SqlDataAdapter(komut);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                aynısipnugetirme();

                if (dataGridView1.Rows.Count > 1)
                {
                    string kayit2 = "SELECT DISTINCT Onay,KesildiMi,SiparisNo,Müşteri,Model,TeslimTarihi,KesildiTarihi,Renk,Palet,Aşama,SiparişTarihi,ToplamM2 From Siparişler where AnaSiparişMi='Evet' AND Renk=@Renk AND Aşama in ('Onay Bekliyor', 'Onaylandı', 'Etiket', 'Palet') AND SiparisNo != @SiparisNo ORDER BY SiparisNo ASC";
                    SqlCommand komut2 = new SqlCommand(kayit2, bgl.baglanti());
                    komut2.Parameters.AddWithValue("@Renk", dataGridView1.Rows[0].Cells["Renk"].Value.ToString());
                    komut2.Parameters.AddWithValue("@SiparisNo", textBox1.Text);
                    SqlDataAdapter da2 = new SqlDataAdapter(komut2);
                    DataTable dt2 = new DataTable();
                    da2.Fill(dt2);
                    dataGridView2.DataSource = dt2;
                }
                aynısipnugetirme2();

            }

        }

        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            aynısipnugetirme();
        }

        private void dataGridView2_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            aynısipnugetirme2();
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                PaletiPresseGonder();
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            checkBox2.Checked = false;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            checkBox1.Checked = false;
        }
    }
}
