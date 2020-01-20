using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Interviewer
{
    public partial class Form2 : Form
    {
        private DataTable dt;
        private string name;
        private int count = 0;

        public Form2(DataTable data, string pers_name)
        {
            InitializeComponent();
            dt = data;
            name = pers_name;
            labelCount.Text = (count+1).ToString() + '/' + dt.Rows.Count.ToString();
            labelQuestion.Text = dt.Rows[count][0].ToString();
            dt.Columns.Add("Answers", typeof(String));
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            if (count < dt.Rows.Count - 1)
            {
                dt.Rows[count][1] = textBox1.Text;
                count++;
                if (count == dt.Rows.Count - 1)
                {
                    btnNext.Text = "Finish";
                }
                
            }
            else
            {
                dt.Rows[count][1] = textBox1.Text;

                using (var workbook = new XLWorkbook())
                {

                    var worksheet = workbook.Worksheets.Add("Test results");
                    for (var i = 0; i < dt.Columns.Count; i++)
                    {
                        worksheet.Cell(1, i + 1).Value = dt.Columns[i].ColumnName;
                    }
                    for (var i = 0; i < dt.Rows.Count; i++)
                    {
                        for (var j = 0; j < dt.Columns.Count; j++)
                        {
                            worksheet.Cell(i + 2, j + 1).Value = dt.Rows[i][j];
                        }
                    }

                    
                    workbook.SaveAs("TestResults_" + name + ".xlsx");
                    MessageBox.Show("Answers saved successfully!","Success", MessageBoxButtons.OK);
                    Application.Exit();
                }

            }

            labelCount.Text = (count + 1).ToString() + '/' + dt.Rows.Count.ToString();
            labelQuestion.Text = dt.Rows[count][0].ToString();
            textBox1.Text = dt.Rows[count][1].ToString();


        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            if (count > 0)
            {
                dt.Rows[count][1] = textBox1.Text;
                count--;
            }
            
            labelCount.Text = (count + 1).ToString() + '/' + dt.Rows.Count.ToString();
            labelQuestion.Text = dt.Rows[count][0].ToString();
            textBox1.Text = dt.Rows[count][1].ToString();
        }

        private void fileToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void quitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string message = "Your answers will not be saved. Continue?";
            string caption = "Are you sure?";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result;

            // Displays the MessageBox.
            result = MessageBox.Show(message, caption, buttons);
            if (result == System.Windows.Forms.DialogResult.Yes)
            {
                // Closes the app.
                Application.Exit();
            }

           


        }
    }
}
