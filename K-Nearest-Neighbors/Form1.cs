using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
//using Microsoft.Office.Interop.Excel;

namespace K_Nearest_Neighbors
{
    public partial class Form1 : Form
    {

        
        //Public Variables
        string fileName;
        int k;
        List<double> x = new List<double>();
        string testData;
        int lastNumber;
        int lastNumber2;
        List<string> classDataList = new List<string>();
        HashSet<string> classDataHash = new HashSet<string>();

        public Form1()
        {
            InitializeComponent();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            k = Convert.ToInt32(numericUpDown1.Value);

            using (OpenFileDialog openFileDialog1 = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel WorkBook 97-2003|*.xls" })
            {
                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    fileName = openFileDialog1.FileName;
                    this.toolStripStatusLabel1.Text = fileName;
                    toolStripStatusLabel1.Text = "Excel file loaded.";
                }
            }

            string PathConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties = \"Excel 12.0 Xml;HDR=YES\"; ";
            OleDbConnection conn = new OleDbConnection(PathConn);
            try
            {
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter("Select * from [" + "Sayfa1" + "$]", conn);
                System.Data.DataTable dt = new System.Data.DataTable();
                myDataAdapter.Fill(dt);
                dataGridView1.DataSource = dt;
                toolStripStatusLabel1.Text = "[TR] File is on view. Ready to run.";
            }
            catch (Exception)
            {
                try
                {
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter("Select * from [" + "Sheet1" + "$]", conn);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    myDataAdapter.Fill(dt);
                    dataGridView1.DataSource = dt;
                    toolStripStatusLabel1.Text = "Excel dosyası yüklenmiştir.";
                }
                catch (Exception)
                {
                    toolStripStatusLabel1.Text = "Wrong language or file format!";
                }
            }
        }

        private double distanceToCenterFunction(double x1Test, double x2Test, double x3Test, double x4Test, double x1Data, double x2Data, double x3Data, double x4Data)
        {
            var result = Math.Sqrt(Math.Pow(x1Test - x1Data, 2) + Math.Pow(x2Test - x2Data, 2) + Math.Pow(x3Test - x3Data, 2) + Math.Pow(x4Test - x4Data, 2));
            return result;
        }

        private void runToolStripMenuItem_Click(object sender, EventArgs e)
        {
            x.Clear();
            string[] myStringList;
            testData = textBoxTestData.Text;
            char seperator = ',';
            myStringList = testData.Split(seperator);

            for (int i = 0; i < myStringList.Length; i++)
            {
                x.Add(Convert.ToDouble(myStringList[i]));
            }
            
            //dataGridView1.Rows[].Cells
            double[] x1Data = new double[dataGridView1.RowCount - 1];
            double[] x2Data = new double[dataGridView1.RowCount - 1];
            double[] x3Data = new double[dataGridView1.RowCount - 1];
            double[] x4Data = new double[dataGridView1.RowCount - 1];
            string[] classData = new string[dataGridView1.RowCount - 1];
            double[] distance = new double[dataGridView1.RowCount - 1];
            
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                x1Data[i] = Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                x2Data[i] = Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                x3Data[i] = Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value);
                x4Data[i] = Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value);
                classData[i] = Convert.ToString(dataGridView1.Rows[i].Cells[5].Value);
                classDataList.Add(Convert.ToString(dataGridView1.Rows[i].Cells[5].Value));
                classDataHash.Add(Convert.ToString(dataGridView1.Rows[i].Cells[5].Value));
                distance[i] = distanceToCenterFunction(x[0], x[1], x[2], x[3], x1Data[i], x2Data[i], x3Data[i], x4Data[i]);
            }
            double[] distanceCopy = new double[dataGridView1.RowCount - 1];
            double[] distanceAscending = new double[dataGridView1.RowCount - 1];
            double[] distanceAscendingIndexNumber = new double[dataGridView1.RowCount - 1];
            double smallest = 9999;
            
            foreach (var item in distance)
            {
                distanceCopy = distance;
            }

            for (int j = 0; j < dataGridView1.RowCount - 1; j++)
            {
                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    if (smallest > distanceCopy[i])
                    {
                        smallest = distanceCopy[i];
                        distanceAscending[j] = smallest;
                        distanceAscendingIndexNumber[j] = i;
                        lastNumber = i;
                    }
                }
                distanceCopy[lastNumber] = 9999;
                smallest = 9999;
            }

            List<string> classList = new List<string>();
            k = Convert.ToInt32(numericUpDown1.Value);
            richTextBoxDistance.Text = null;
            richTextBoxIndex.Text = null;
            richTextBoxClass.Text = null;

            for (int i = 0; i < k; i++)
            {
                richTextBoxDistance.AppendText(distanceAscending[i].ToString() + Environment.NewLine);
                richTextBoxIndex.AppendText(distanceAscendingIndexNumber[i].ToString() + Environment.NewLine);
                richTextBoxClass.AppendText(classData[Convert.ToInt32(distanceAscendingIndexNumber[i])] + Environment.NewLine);
                classList.Add(classData[Convert.ToInt32(distanceAscendingIndexNumber[i])]);
            }
            
            string[] classTypeList = classDataHash.ToArray();
            int[] numberList = new int[classTypeList.Length];
            for (int j = 0; j < classTypeList.Length; j++)
            {
                for (int i = 0; i < classList.Count; i++)
                {
                    string item1 = classTypeList[j].ToString();
                    string item2 = classList[i].ToString();
                    if (item1 == item2)
                    {
                        numberList[j] = numberList[j] + 1;
                    }
                }
            }
            
            int biggest = 0;
            
            for (int a = 0; a < numberList.Length; a++)
            {
                if (biggest < numberList[a])
                {
                    biggest = numberList[a];
                    lastNumber2 = a;
                }
            }
            //MessageBox.Show("The class of test data is:" + " " + classTypeList[lastNumber2]);
            toolStripStatusLabel1.Text = "Analiz Tamamlandı:" + " " + classTypeList[lastNumber2];
        }
    }
}
