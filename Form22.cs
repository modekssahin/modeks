using FirebirdSql.Data.FirebirdClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Modeks
{
    public partial class Form22 : Form
    {
        public Form22()
        {
            InitializeComponent();
        }

        private void Form22_Load(object sender, EventArgs e)
        {

        }
        int Musteri_ID = 100000;
        private void musteri_ıd_getir()
        {
            try
            {
                var connectionString = @"User ID=SYSDBA;Password=masterkey;Database=localhost:C:\Program Files (x86)\BSR\VERESIYEDATA.FDB ;Charset=WIN1254;";
                FbConnection fbcnn = new FbConnection(connectionString);
                string sql = "select * from MUSTERILER";
                fbcnn.Open();
                FbCommand command = new FbCommand(sql, fbcnn);
                FbDataReader reader = command.ExecuteReader();
                StringBuilder sb = new StringBuilder();
                while (reader.Read())
                {
                    if (Musteri_ID < Convert.ToInt32(reader["ID"].ToString()))
                    {
                        Musteri_ID = Convert.ToInt32(reader["ID"]);
                    }

                }
            }
            catch (Exception)
            {

            }

        }
        private void MusteriEkle()
        {
            if (textBox1.Text != "" && textBox2.Text != "" && textBox3.Text != "")
            {
                musteri_ıd_getir();
                var connectionString = @"User ID=SYSDBA;Password=masterkey;Database=localhost:C:\Program Files (x86)\BSR\VERESIYEDATA.FDB ;Charset=WIN1254;";
                using (FbConnection fbcnn = new FbConnection(connectionString))
                {
                    fbcnn.Open();

                    string kayit = "insert into MUSTERILER(ID,ADI_SOYADI,TELEFON,ADRES) values (@ID,@ADI_SOYADI,@TELEFON,@ADRES)";
                    using (FbCommand komut = new FbCommand(kayit, fbcnn))
                    {
                        komut.Parameters.AddWithValue("@ID", (Musteri_ID + 1));
                        komut.Parameters.AddWithValue("@ADI_SOYADI", textBox1.Text);
                        komut.Parameters.AddWithValue("@TELEFON", textBox2.Text);

                        // BLOB için akış oluşturup veriyi yazın
                        using (FbTransaction transaction = fbcnn.BeginTransaction())
                        {
                            komut.Transaction = transaction;

                            byte[] adresData = System.Text.Encoding.Default.GetBytes(textBox3.Text);
                            FbParameter blobParameter = new FbParameter("@ADRES", FbDbType.Binary);
                            blobParameter.Value = adresData;
                            komut.Parameters.Add(blobParameter);

                            komut.ExecuteNonQuery();

                            transaction.Commit();
                        }
                    }
                    fbcnn.Close();
                    MessageBox.Show("Müşteri Başarıyla Eklenmiştir.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    textBox1.Text = "";
                    textBox2.Text = "";
                    textBox3.Text = "";
                }
            }
            else
            {
                MessageBox.Show("Lütfen müşteri ekleyebilmek için tüm bilgileri giriniz!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                MusteriEkle();
                Form10 frm = (Form10)Application.OpenForms["Form10"];
                frm.müsteri_cek();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox1.Text = textBox1.Text.ToUpper();
            textBox1.SelectionStart = textBox1.Text.Length;
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                MusteriEkle();
            }
        }

        private void Form22_FormClosing(object sender, FormClosingEventArgs e)
        {

        }
    }
}