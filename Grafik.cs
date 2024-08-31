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
using System.Windows.Forms.DataVisualization.Charting;

namespace Modeks
{
    public partial class Grafik : Form
    {
        public Grafik()
        {
            InitializeComponent();
        }

        private sqlsinif bgl = new sqlsinif();

        private void Grafik_Load(object sender, EventArgs e)
        {
            using (System.Data.SqlClient.SqlConnection connection = bgl.baglanti())
            {
                string sorgu = "SELECT Renk, ToplamM2, PressHazirM2, PaletlenecekM2, OnayBekleyenM2, PaletNumarasi FROM RenkDurum ORDER BY PressHazirM2 ASC";

                System.Data.SqlClient.SqlDataAdapter adapter = new System.Data.SqlClient.SqlDataAdapter(sorgu, connection);
                System.Data.DataTable dataTable = new System.Data.DataTable();
                adapter.Fill(dataTable);

                chart1.Series.Clear();

                System.Windows.Forms.DataVisualization.Charting.Series seriesPressHazir = new System.Windows.Forms.DataVisualization.Charting.Series("Press Hazır");
                System.Windows.Forms.DataVisualization.Charting.Series seriesPaletlenecek = new System.Windows.Forms.DataVisualization.Charting.Series("Paletlenecek");
                seriesPressHazir.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.StackedBar;
                seriesPaletlenecek.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.StackedBar;

                foreach (System.Data.DataRow row in dataTable.Rows)
                {
                    string color = row["Renk"].ToString();
                    double totalM2 = Convert.ToDouble(row["ToplamM2"]);
                    double pressReadyM2 = Convert.ToDouble(row["PressHazirM2"]);
                    double palletM2 = Convert.ToDouble(row["PaletlenecekM2"]);
                    string palletNumber = row["PaletNumarasi"].ToString();

                    double pressHazirPercentage = totalM2 > 0 ? (pressReadyM2 / totalM2) * 100 : 0;
                    double paletlenecekPercentage = totalM2 > 0 ? (palletM2 / totalM2) * 100 : 0;

                    seriesPressHazir.Points.AddXY(color, pressHazirPercentage);
                    seriesPaletlenecek.Points.AddXY(color, paletlenecekPercentage);

                    var pressHazirPoint = seriesPressHazir.Points[seriesPressHazir.Points.Count - 1];
                    var paletlenecekPoint = seriesPaletlenecek.Points[seriesPaletlenecek.Points.Count - 1];

                    pressHazirPoint.Label = $"{pressHazirPercentage:F1}%";
                    paletlenecekPoint.Label = $"{paletlenecekPercentage:F1}%";

                    if (palletNumber == "Numara Yok")
                    {
                        pressHazirPoint.LabelForeColor = Color.Red;
                        paletlenecekPoint.LabelForeColor = Color.Red;
                        pressHazirPoint.Label = $"{pressHazirPercentage:F2}%";
                        paletlenecekPoint.Label = $"{paletlenecekPercentage:F2}%";
                    }
                    else
                    {
                        pressHazirPoint.Label = $"{pressHazirPercentage:F1}%\nPalet: {palletNumber}";
                        paletlenecekPoint.Label = $"{paletlenecekPercentage:F1}%\nPalet: {palletNumber}";
                    }
                }

                chart1.Series.Add(seriesPressHazir);
                chart1.Series.Add(seriesPaletlenecek);

                chart1.ChartAreas[0].AxisX.Title = "Yüzde";
                chart1.ChartAreas[0].AxisY.Title = "Renk";
                chart1.ChartAreas[0].AxisX.LabelStyle.Format = "{0}%";
                chart1.ChartAreas[0].AxisY.Interval = 1; 
                chart1.Legends[0].DockedToChartArea = chart1.ChartAreas[0].Name;
                chart1.Legends[0].Alignment = System.Drawing.StringAlignment.Center;
            }
        }
    }
}

