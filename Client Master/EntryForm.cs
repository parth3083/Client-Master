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
            string query = "INSERT INTO ClientMaster (Company, PCS, DateOfAppointmentLetter, DateOfMOU, Charges, CDSLISIN, CDSLTripartiteDate, CDSLRemark, NSDLISIN, NSDLTripartiteDate, NSDLRemark, GST, ContactPerson, MobileNo, Email, ConnectivityDoc, AuthorityResolution, Signature) " +
                       "VALUES (@Company, @PCS, @DateOfAppointmentLetter, @DateOfMOU, @Charges, @CDSLISIN, @CDSLTripartiteDate, @CDSLRemark, @NSDLISIN, @NSDLTripartiteDate, @NSDLRemark, @GST, @ContactPerson, @MobileNo, @Email, @ConnectivityDoc, @AuthorityResolution, @Signature)";
            using(SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlCommand command= new SqlCommand(query, conn);
                command.Parameters.AddWithValue("@Company", Company_name.Text);
                command.Parameters.AddWithValue("@PCS", PCS_value.Text);
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
                command.Parameters.AddWithValue("@ConnectivityDoc", Connectivity_doc.Checked ? 1 : 0);
                command.Parameters.AddWithValue("@AuthorityResolution", authority_resolution_yes.Checked ? 1 : 0);
                command.Parameters.AddWithValue("@Signature", signature_yes.Checked ? 1 : 0);
                conn.Open();
                command.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Data saved successfully");

            }
        }
    }
}
