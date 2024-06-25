using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Client_Master
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            int numberOfEntries;
            if (int.TryParse(textBox1.Text, out numberOfEntries) && numberOfEntries > 0)
            {
                for (int i = 0; i < numberOfEntries; i++)
                {
                    var entryForm = new EntryForm();
                    entryForm.ShowDialog();
                }

            }
            else
            {
                MessageBox.Show("Please enter a valid number of entries.");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var updateForm = new UpdateForm();
            updateForm.ShowDialog();
        }
    }
}
