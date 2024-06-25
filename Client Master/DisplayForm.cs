using DocumentFormat.OpenXml.Office2010.Excel;
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

namespace Client_Master
{
    public partial class DisplayForm : Form
    {
        private int _id;
        public DisplayForm(int id)
        {
            _id = id;
            InitializeComponent();
            LoadData();
        }
        private void LoadData()
        {
            string connectionString = "Data Source=PARTH\\SQLEXPRESS;Initial Catalog=ClientMaster;Integrated Security=True;Encrypt=False";
            string query = "SELECT * FROM ClientMaster WHERE Id = @Id";
            using(SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@ID", _id);
                connection.Open();
                SqlDataReader reader= cmd.ExecuteReader();
                if (reader.Read())
                {
                    company.Text = reader["Company"].ToString();
                    PCS.Text = reader["PCS"].ToString();
                    company_email.Text = reader["CompanyEmail"].ToString();
                    PCS_email.Text = reader["PCSEmail"].ToString();
                    date_of_appointment_letter.Value = Convert.ToDateTime(reader["DateOfAppointmentLetter"]);
                    Date_of_MOU.Value = Convert.ToDateTime(reader["DateOfMOU"]);
                    charges.Text = reader["Charges"].ToString();
                    CSDL_ISIN.Text = reader["CDSLISIN"].ToString();
                    CSDL.Value = Convert.ToDateTime(reader["CDSLTripartiteDate"]);
                    CSDL_remark.Text = reader["CDSLRemark"].ToString();
                    NSDL_ISIN.Text = reader["NSDLISIN"].ToString();
                    NSDL.Value = Convert.ToDateTime(reader["NSDLTripartiteDate"]);
                    NSDL_remark.Text = reader["NSDLRemark"].ToString();
                    GST.Text = reader["GST"].ToString();
                    contact_person.Text = reader["ContactPerson"].ToString();
                    mobile.Text = reader["MobileNo"].ToString();
                    email.Text = reader["Email"].ToString();
                    creativity_doc_yes.Checked = (bool)reader["ConnectivityDoc"];
                    authority_resolution_yes.Checked = (bool)reader["AuthorityResolution"];
                    signature_yes.Checked = (bool)reader["Signature"];
                }
            }
        }

        private void DisplayForm_Load(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label23_Click(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void label18_Click(object sender, EventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void label21_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label22_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string connectionString = "Data Source=PARTH\\SQLEXPRESS;Initial Catalog=ClientMaster;Integrated Security=True;Encrypt=False";
            string query = "UPDATE ClientMaster SET Company = @Company,CompanyEmail=@CompanyEmail,PCSEmail=@PCSEmail, PCS = @PCS, DateOfAppointmentLetter = @DateOfAppointmentLetter, DateOfMOU = @DateOfMOU, Charges = @Charges, CDSLISIN = @CDSLISIN, CDSLTripartiteDate = @CDSLTripartiteDate, CDSLRemark = @CDSLRemark, NSDLISIN = @NSDLISIN, NSDLTripartiteDate = @NSDLTripartiteDate, NSDLRemark = @NSDLRemark, GST = @GST, ContactPerson = @ContactPerson, MobileNo = @MobileNo, Email = @Email, ConnectivityDoc = @ConnectivityDoc, AuthorityResolution = @AuthorityResolution, Signature = @Signature WHERE Id = @Id";
            try
            {
                using(SqlConnection connection =new SqlConnection(connectionString))
                {
                    SqlCommand command =new SqlCommand(query, connection);
                    command.Parameters.AddWithValue("@Company", company.Text);
                    command.Parameters.AddWithValue("@PCS", PCS.Text);
                    command.Parameters.AddWithValue("@CompanyEmail", company_email.Text);
                    command.Parameters.AddWithValue("@PCSEmail", PCS_email.Text);
                    command.Parameters.AddWithValue("@DateOfAppointmentLetter", date_of_appointment_letter.Value);
                    command.Parameters.AddWithValue("@DateOfMOU", Date_of_MOU.Value);
                    command.Parameters.AddWithValue("@Charges", charges.Text);
                    command.Parameters.AddWithValue("@CDSLISIN", CSDL_ISIN.Text);
                    command.Parameters.AddWithValue("@CDSLTripartiteDate", CSDL.Value);
                    command.Parameters.AddWithValue("@CDSLRemark", CSDL_remark.Text);
                    command.Parameters.AddWithValue("@NSDLISIN", NSDL_ISIN.Text);
                    command.Parameters.AddWithValue("@NSDLTripartiteDate", NSDL.Value);
                    command.Parameters.AddWithValue("@NSDLRemark", NSDL_remark.Text);
                    command.Parameters.AddWithValue("@GST", GST.Text);
                    command.Parameters.AddWithValue("@ContactPerson", contact_person.Text);
                    command.Parameters.AddWithValue("@MobileNo", mobile.Text);
                    command.Parameters.AddWithValue("@Email", email.Text);
                    command.Parameters.AddWithValue("@ConnectivityDoc", creativity_doc_yes.Checked ? 1 : 0);
                    command.Parameters.AddWithValue("@AuthorityResolution", authority_resolution_yes.Checked ? 1 : 0);
                    command.Parameters.AddWithValue("@Signature", signature_yes.Checked ? 1 : 0);
                    command.Parameters.AddWithValue("@Id", _id);

                    connection.Open();
                    command.ExecuteNonQuery();
                    connection.Close();

                    MessageBox.Show("Data updated successfully!");
                    this.Close();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
    }
}
