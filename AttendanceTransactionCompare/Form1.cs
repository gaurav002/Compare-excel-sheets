using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AttendanceTransactionCompare
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
         

        }


        private void button2_Click(object sender, System.EventArgs e)
        {
            throw new System.NotImplementedException();
        }

        private void button3_Click(object sender, System.EventArgs e)
        {
            throw new System.NotImplementedException();
        }

        private void button4_Click(object sender, System.EventArgs e)
        {
            throw new System.NotImplementedException();
        }

        private void button5_Click(object sender, System.EventArgs e)
        {
            throw new System.NotImplementedException();
        }

        private void btnAuthorizationCompare_Click(object sender, System.EventArgs e)
        {
            throw new System.NotImplementedException();
        }

        private void button6_Click(object sender, System.EventArgs e)
        {
            throw new System.NotImplementedException();
        }

        private void button7_Click(object sender, System.EventArgs e)
        {
            throw new System.NotImplementedException();
        }

        private void button8_Click(object sender, System.EventArgs e)
        {
            throw new System.NotImplementedException();
        }

        private void button9_Click(object sender, System.EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFilePath.Text.Trim()))
            {
                MessageBox.Show("Please select the file..", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            ProcessExcel process = new ProcessExcel();
            //process.ProcessPSSData(txtFilePath.Text.Trim());
            //process.ProcessPOCData(txtFilePath.Text.Trim());
            //process.ProcessPOCPSSCommonData(txtFilePath.Text.Trim());
            process.ProcessPOCPSSDuplicateData(txtFilePath.Text.Trim());
            txtStatus.AppendText("Process Completed.." + Environment.NewLine);

         
        }

        private void button10_Click(object sender, System.EventArgs e)
        {
            throw new System.NotImplementedException();
        }

        private void button11_Click(object sender, System.EventArgs e)
        {
            var result = OFDialog.ShowDialog();
            txtFilePath.Text = OFDialog.FileName;           
        }
    }
}
