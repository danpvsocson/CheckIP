using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using OfficeOpenXml;
using ExcelDataReader;
using Z.Dapper.Plus;
using System.IO;
using System.Diagnostics;
using WPFCustomMessageBox;

namespace CheckIP
{
    public partial class Form1 : Form
    {
        private DataTableCollection tables;
        private DataTable excelDataTable;
        private string connectionString = "Server=CAD001\\WEB;Database=MTH;User=sa;Password=abc123";
        //private readonly string outputFilePath = "C:\\User\\dan\\Desktop\\output.txt";
        private readonly string desktopPath;
        private readonly string outputFilePath;
        public Form1()
        {
            InitializeComponent();
            // Initialize desktopPath in the constructor
            desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            // Initialize outputFilePath using desktopPath
            outputFilePath = Path.Combine(desktopPath, "output.txt");
        }

        private void find_Click(object sender, EventArgs e)
        {

            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls" })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    txtPath.Text = ofd.FileName;
                    using (var stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = true
                                }
                            });
                            tables = result.Tables;
                            comboBox1.Items.Clear();
                            foreach (DataTable table in tables)
                                comboBox1.Items.Add(table.TableName);
                        }
                    }
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null)
            {
                DataTable dt = tables[comboBox1.SelectedItem.ToString()];

                // Remove empty columns
                List<DataColumn> columnsToRemove = new List<DataColumn>();
                foreach (DataColumn column in dt.Columns)
                {
                    bool isColumnEmpty = true;
                    foreach (DataRow row in dt.Rows)
                    {
                        if (!string.IsNullOrWhiteSpace(row[column.ColumnName].ToString()))
                        {
                            isColumnEmpty = false;
                            break;
                        }
                    }
                    if (isColumnEmpty)
                    {
                        columnsToRemove.Add(column);
                    }
                }

                foreach (DataColumn columnToRemove in columnsToRemove)
                {
                    dt.Columns.Remove(columnToRemove);
                }

                dataGridView1.DataSource = dt;
            }
        }

        private void btnCheck_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null)
            {
                DataTable dt = tables[comboBox1.SelectedItem.ToString()];
                progressBar.Minimum = 0;
                progressBar.Maximum = dt.Rows.Count;
                progressBar.Step = 1;
                progressBar.Value = 0;
                using (StreamWriter writer = new StreamWriter(outputFilePath))
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        string ip = row["ip"].ToString();
                        bool isReachable = CheckIP(ip);
                        string status = isReachable ? "Online" : "Offline";
                        writer.WriteLine($"IP: {ip}, Trang Thai: {status}");

                        progressBar.PerformStep();
                        int percent = (progressBar.Value * 100) / progressBar.Maximum;
                        progressBar.CreateGraphics().DrawString(percent.ToString() + "%", new Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressBar.Width / 2 - 10, progressBar.Height / 2 - 7));
                    }
                }
                //CustomMessageBox.ShowOKCancel(
                //"Are you sure you want to eject the nuclear fuel rods?",
                //"Confirm Fuel Ejection",
                //"Eject Fuel Rods",
                //"Don't do it!");
                DialogResult result = MessageBox.Show("File check đã được lưu trên Desktop. Mở File Ngay", "Thông báo", MessageBoxButtons.OKCancel);
                if (result == DialogResult.OK)
                {

                    // Mở file output với ứng dụng mặc định
                    Process.Start(outputFilePath);

                    // Đóng ứng dụng
                    Application.Exit();
                }
                else if (result == DialogResult.Cancel)
                {

                    // Đóng ứng dụng nếu người dùng nhấn Cancel
                    Application.Exit();
                }
            }
        }
        private bool CheckIP(string ip)
        {
            try
            {
                System.Net.NetworkInformation.Ping ping = new System.Net.NetworkInformation.Ping();
                var reply = ping.Send(ip);

                // Đặt giá trị isReachable là true khi Status là Success hoặc TimedOut
                return reply.Status == System.Net.NetworkInformation.IPStatus.Success || reply.Status == System.Net.NetworkInformation.IPStatus.TimedOut;
            }
            catch (Exception)
            {
                // Xử lý lỗi (nếu cần thiết)
                return false;
            }
        }
    }
}
