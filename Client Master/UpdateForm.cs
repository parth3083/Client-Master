using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Client_Master
{
    public partial class UpdateForm : Form
    {
        public UpdateForm()
        {
            InitializeComponent();
        }

        private void UpdateForm_Load(object sender, EventArgs e)
        {
            string connectionString = "Data Source=PARTH\\SQLEXPRESS;Initial Catalog=ClientMaster;Integrated Security=True;Encrypt=False";
            string query = "select * from ClientMaster";
            using(SqlConnection connection =new SqlConnection(connectionString))
            {
                SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
                DataTable dt = new DataTable();
                adapter.Fill(dt);
                dataGridView1.DataSource = dt;

            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            int id;
            if (int.TryParse(dataID.Text, out id))
            {
                //UpdateForm updateForm = new UpdateForm(id);
                //updateForm.Show();
                DisplayForm displayForm = new DisplayForm(id);
                displayForm.Show();
            }
            else
            {
                MessageBox.Show("Please enter a valid ID.");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string connectionString = "Data Source=PARTH\\SQLEXPRESS;Initial Catalog=ClientMaster;Integrated Security=True;Encrypt=False";
            string query = "select * from ClientMaster";
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    SqlDataAdapter sda = new SqlDataAdapter(query, connection);
                    DataTable dt = new DataTable();
                    sda.Fill(dt);
                    using (XLWorkbook workbook = new XLWorkbook())
                    {
                        workbook.Worksheets.Add(dt, "Client Master Data");
                        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                        string filePath = Path.Combine(desktopPath, "ClientMasterData.xlsx");
                        workbook.SaveAs(filePath);
                        MessageBox.Show("Excel sheet saved to Desktop");
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
