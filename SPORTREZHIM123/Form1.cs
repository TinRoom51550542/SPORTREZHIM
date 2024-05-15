using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using System.Windows.Forms.DataVisualization.Charting;



namespace SPORTREZHIM123
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btn_OpenFile(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            FileInfo fileInfo = new FileInfo(@"C:\Users\Евгения\OneDrive\Рабочий стол\StatisticOfRunning.xlsx");
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                for (int i = 1; i <= worksheet.Dimension.Columns; i++)
                {
                    dataGridView1.Columns.Add("Column" + i, worksheet.Cells[1, i].Value.ToString());
                }

                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    object[] values = new object[worksheet.Dimension.Columns];
                    for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                    {
                        values[col - 1] = worksheet.Cells[row, col].Value;
                    }
                    dataGridView1.Rows.Add(values);
                }
            }

        }

        private void btnSumKm_Click(object sender, EventArgs e)
        {

        }

        private void btnForecast_Click(object sender, EventArgs e)
        {

        }

        private void btnPaintGraphics_Click(object sender, EventArgs e)
        {
            chart1.Series.Clear();
            chart1.ChartAreas.Clear();

            FileInfo fileInfo = new FileInfo(@"C:\Users\PC\Desktop\WindowsFormsApp1\1.xlsx");
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                chart1.ChartAreas.Add(new ChartArea("Area1"));

                Series series1 = new Series
                {
                    Name = "Длительность бега",
                    ChartType = SeriesChartType.Line
                };

                Series series2 = new Series
                {
                    Name = "Скорость",
                    ChartType = SeriesChartType.Line
                };

                chart1.Series.Add(series1);

                chart1.Series["Длительность бега"].XValueType = ChartValueType.String;

                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    string date = worksheet.Cells[row, 1].Text;
                    double runDuration = Convert.ToDouble(worksheet.Cells[row, 3].Text);
                    double speed = Convert.ToDouble(worksheet.Cells[row, 5].Text.Replace(',', '.'));

                    chart1.Series["Длительность бега"].Points.AddXY(date, runDuration);
                }
            }
        }
    }
}
