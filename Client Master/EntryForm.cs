using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Client_Master
{
    public partial class EntryForm : Form
    {
        public EntryForm()
        {
            InitializeComponent();
        }

        private void EntryForm_Load(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void NSDLISIN_TextChanged(object sender, EventArgs e)
        {

        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {

        }

        private void submit_Click(object sender, EventArgs e)
        {
            string connectionString = "Data Source=PARTH\\SQLEXPRESS;Initial Catalog=ClientMaster;Integrated Security=True;Encrypt=False";
            string query = "INSERT INTO ClientMaster (Company, PCS, CompanyEmail, PCSEmail, DateOfAppointmentLetter, DateOfMOU, Charges, CDSLISIN, CDSLTripartiteDate, CDSLRemark, NSDLISIN, NSDLTripartiteDate, NSDLRemark, GST, ContactPerson, MobileNo, Email, ConnectivityDoc, AuthorityResolution, Signature) " +
                        "VALUES (@Company, @PCS, @CompanyEmail, @PCSEmail, @DateOfAppointmentLetter, @DateOfMOU, @Charges, @CDSLISIN, @CDSLTripartiteDate, @CDSLRemark, @NSDLISIN, @NSDLTripartiteDate, @NSDLRemark, @GST, @ContactPerson, @MobileNo, @Email, @ConnectivityDoc, @AuthorityResolution, @Signature)";
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    SqlCommand command = new SqlCommand(query, conn);
                    command.Parameters.AddWithValue("@Company", Company_name.Text);
                    command.Parameters.AddWithValue("@PCS", PCS_value.Text);
                    command.Parameters.AddWithValue("@CompanyEmail", company_email.Text);
                    command.Parameters.AddWithValue("@PCSEmail", PCS_email.Text);
                    command.Parameters.AddWithValue("@DateOfAppointmentLetter", Date_of_appiontment_letter.Value);
                    command.Parameters.AddWithValue("@DateOfMOU", Date_of_MOU.Value);
                    command.Parameters.AddWithValue("@Charges", Charges.Text);
                    command.Parameters.AddWithValue("@CDSLISIN", CSDLISIN.Text);
                    command.Parameters.AddWithValue("@CDSLTripartiteDate", CSDL.Value);
                    command.Parameters.AddWithValue("@CDSLRemark", CSDL_Remark.Text);
                    command.Parameters.AddWithValue("@NSDLISIN", NSDLISIN.Text);
                    command.Parameters.AddWithValue("@NSDLTripartiteDate", NSDL.Value);
                    command.Parameters.AddWithValue("@NSDLRemark", NSDL_Remark.Text);
                    command.Parameters.AddWithValue("@GST", GST.Text);
                    command.Parameters.AddWithValue("@ContactPerson", Contact_Person.Text);
                    command.Parameters.AddWithValue("@MobileNo", Mobile.Text);
                    command.Parameters.AddWithValue("@Email", Email.Text);
                    command.Parameters.AddWithValue("@ConnectivityDoc", Connectivity_doc_yes.Checked ? 1 : 0);
                    command.Parameters.AddWithValue("@AuthorityResolution", authority_resolution_yes.Checked ? 1 : 0);
                    command.Parameters.AddWithValue("@Signature", signature_yes.Checked ? 1 : 0);
                    conn.Open();
                    command.ExecuteNonQuery();
                    conn.Close();
                    MessageBox.Show("Data saved successfully");

                }
                this.Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void get_excel_Click(object sender, EventArgs e)
        {
            string connectionString = "Data Source=PARTH\\SQLEXPRESS;Initial Catalog=ClientMaster;Integrated Security=True;Encrypt=False";
            string query = "select * from ClientMaster";
            try
            {
                using(SqlConnection connection =new SqlConnection(connectionString))
                {
                    SqlDataAdapter sda =new SqlDataAdapter(query, connection);
                    DataTable dt = new DataTable();
                    sda.Fill(dt);
                    using(XLWorkbook workbook =new XLWorkbook())
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
