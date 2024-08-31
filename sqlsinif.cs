using System.Data.SqlClient;

namespace Modeks
{
    public class sqlsinif
    {
        public SqlConnection baglanti()
        {
            SqlConnection baglan = new SqlConnection("Data Source=78.108.246.74;Initial Catalog=Modeks_2022;User ID=modeksadmin;Password=8659745Modeks;Encrypt=True;TrustServerCertificate=true;"); 
            baglan.Open();
            return baglan;
        }
        public SqlConnection baglanti_eski()
        {
            SqlConnection baglan = new SqlConnection("Data Source=78.108.246.74;Initial Catalog=Modeks_Eski;User ID=modeksadmin;Password=8659745Modeks;Encrypt=True;TrustServerCertificate=true;");
            baglan.Open();
            return baglan;
        }
    }
}