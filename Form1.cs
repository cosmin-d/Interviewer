using ExcelDataReader;
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

namespace Interviewer
{
   
    public partial class Form1 : Form
    {
        private DataTable dt;
        private string name;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void cboSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
             dt = tableCollection[cboSheet.SelectedItem.ToString()];
            dataGridView1.DataSource = dt;
            if (dt != null) 
                if (dt.Rows.Count > 0) { btnNext.Enabled = true; }
            else btnNext.Enabled = false;
        }

        DataTableCollection tableCollection;

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx;*.xls" })
            {
                if(openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFilename.Text = openFileDialog.FileName;
                    using (var stream = File.Open(openFileDialog.FileName,FileMode.Open, FileAccess.Read))
                    {
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });
                            tableCollection = result.Tables;
                            cboSheet.Items.Clear();
                            foreach (DataTable table in tableCollection)
                                cboSheet.Items.Add(table.TableName);
                        }
                    }
                }
            }
                
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("Pressed", "Test Button", MessageBoxButtons.OK);
            name = tb_name.Text;
            Form2 frm2 = new Form2(dt,name);
            frm2.Show();
            this.Hide();
        }

        private void quitToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
             Application.Exit();
            
        }
    }
}
