using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;

namespace Modeks
{
    public partial class Kullanicilar : Form
    {
        sqlsinif db = new sqlsinif();
        private List<string> permissions = new List<string>();

        public Kullanicilar()
        {
            InitializeComponent();
        }

        private void Kullanicilar_Load(object sender, EventArgs e)
        {
            LoadUsers();
            LoadPermissions();
        }

        private void LoadUsers()
        {
            userList.Items.Clear();
            SqlCommand command = new SqlCommand("SELECT [id], [kullanıcıadı] FROM [dbo].[Giriş]", db.baglanti());
            SqlDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                userList.Items.Add($"{reader["id"]} - {reader["kullanıcıadı"]}");
            }
            db.baglanti().Close();
        }

        private void LoadPermissions()
        {
            
            permissionComboBox.Items.Clear();
            permissionComboBox.Items.Add("Yönetici");
        }

        private void UserList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (userList.SelectedItem != null)
            {
                string selectedUser = userList.SelectedItem.ToString();
                string[] userInfo = selectedUser.Split(' ');
                string id = userInfo[0];

                SqlCommand command = new SqlCommand("SELECT [id], [kullanıcıadı], [şifre], [yetki] FROM [dbo].[Giriş] WHERE id = @id", db.baglanti());
                command.Parameters.AddWithValue("@id", id);
                SqlDataReader reader = command.ExecuteReader();

                if (reader.Read())
                {
                    idTextBox.Text = reader["id"].ToString();
                    usernameTextBox.Text = reader["kullanıcıadı"].ToString();
                    passwordTextBox.Text = reader["şifre"].ToString();
                    permissionComboBox.SelectedItem = reader["yetki"].ToString();
                }
                db.baglanti().Close();
            }
        }

        private void AddButton_Click(object sender, EventArgs e)
        {
            SqlCommand command = new SqlCommand("INSERT INTO [dbo].[Giriş] ([kullanıcıadı], [şifre], [yetki]) VALUES (@username, @password, @permission)", db.baglanti());
            command.Parameters.AddWithValue("@username", usernameTextBox.Text);
            command.Parameters.AddWithValue("@password", passwordTextBox.Text);
            command.Parameters.AddWithValue("@permission", permissionComboBox.SelectedItem);
            command.ExecuteNonQuery();
            db.baglanti().Close();
            LoadUsers();
        }

        private void UpdateButton_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(idTextBox.Text))
            {
                SqlCommand command = new SqlCommand("UPDATE [dbo].[Giriş] SET [kullanıcıadı] = @username, [şifre] = @password, [yetki] = @permission WHERE id = @id", db.baglanti());
                command.Parameters.AddWithValue("@id", idTextBox.Text);
                command.Parameters.AddWithValue("@username", usernameTextBox.Text);
                command.Parameters.AddWithValue("@password", passwordTextBox.Text);
                command.Parameters.AddWithValue("@permission", permissionComboBox.SelectedItem);
                command.ExecuteNonQuery();
                db.baglanti().Close();
                LoadUsers();
            }
        }

        private void DeleteButton_Click(object sender, EventArgs e)
        {
            if (userList.SelectedItem != null)
            {
                string selectedUser = userList.SelectedItem.ToString();
                string id = selectedUser.Split(' ')[0];

                SqlCommand command = new SqlCommand("DELETE FROM [dbo].[Giriş] WHERE id = @id", db.baglanti());
                command.Parameters.AddWithValue("@id", id);
                command.ExecuteNonQuery();
                db.baglanti().Close();
                LoadUsers();
            }
        }

        private void RefreshButton_Click(object sender, EventArgs e)
        {
            LoadUsers();
        }
    }
}
