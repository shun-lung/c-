using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
//using Excel = Microsoft.Office.Interop.Excel;

using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;


namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }







        private void button1_Click(object sender, EventArgs e)

        {
           // openFileDialog1.Filter = "Excel |*.xls";
          //  openFileDialog1.Title = "open excel file";

            //    //openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);



            //    if (openFileDialog1.ShowDialog() == DialogResult.OK)
            //    {
            //        System.Windows.Forms.WebBrowser wb = new System.Windows.Forms.WebBrowser();
            //       wb.Navigate(openFileDialog1.FileName);
            //        text_Excel.Text = openFileDialog1.FileName;

            //    }

            text_Excel.Text = "C:\\Users\\vita\\Documents\\5.winform\\123.xls";

            int P_int_count = 0;
            string P_str_Line, P_str_content = "";
            List<string> P_list = new List<string>();
            StreamReader sreader = new StreamReader(text_Excel.Text, Encoding.Default);

            while ((P_str_Line = sreader.ReadLine()) != null)
            {
                P_list.Add(P_str_Line);
                P_int_count++;
            }


            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            object missing = System.Reflection.Missing.Value;
            Workbook workbook = excel.Application.Workbooks.Open(text_Excel.Text, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
            Worksheet newWorksheet = (Worksheet)workbook.Worksheets.Add(missing, missing, missing, missing);

            excel.Application.DisplayAlerts = false;
            for (int i = 0; i < P_list.Count; i++)
            {
                P_str_content = P_list[i];
                if (Regex.IsMatch(P_str_content, "^[0-9]*[1-9][0-9]*$"))
                    newWorksheet.Cells[i + 1, 1] = Convert.ToDecimal(P_str_content).ToString("¥00.00");
                else
                    newWorksheet.Cells[i + 1, 1] = P_str_content;
            }

            workbook.Save();
            workbook.Close(false, missing, missing);
            MessageBox.Show("save ok ");

        }


       




        private void btn_upload_Click_1(object sender, EventArgs e)
        {
        string file = "";   //variable for the Excel File Location
        System.Data.DataTable dt = new System.Data.DataTable();   //container for our excel data

        DataRow row;
        DialogResult result = openFileDialog1.ShowDialog();  // Show the dialog.

            

            if (result == DialogResult.OK)   // Check if Result == "OK".
            {
                file = openFileDialog1.FileName; //get the filename with the location of the file
                try

                {
                    //Create Object for Microsoft.Office.Interop.Excel that will be use to read excel file

                    Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

                    Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(file);

                    Microsoft.Office.Interop.Excel._Worksheet excelWorksheet = excelWorkbook.Sheets[1];

                    Microsoft.Office.Interop.Excel.Range excelRange = excelWorksheet.UsedRange;

                   // object[,] data1 = excelRange.Value2;
                    object[,] data = (object[,])excelRange.Value2;


                    int rowCount = excelRange.Rows.Count;  //get row count of excel data

                    int colCount = excelRange.Columns.Count; // get column count of excel data

                    //Get the first Column of excel file which is the Column Name


                    for (int i = 1; i <= rowCount; i++)
                    {

                        //dt.Columns.Add(excelRange.Cells[i, j].Value2.ToString());
                        //
                        var Column = new DataColumn();
                        Column.DataType = System.Type.GetType("System.String");
                        Column.ColumnName = i.ToString();
                        dt.Columns.Add(Column);
                        for (int j = 1; j <= colCount; j++)
                        {
                            text_Excel.Text = j.ToString();

                        }
                        break;
                    }
                    //Get Row Data of Excel              
                    int rowCounter;  //This variable is used for row index number
                    for (int i = 2; i <= rowCount; i++) //Loop for available row of excel data
                    {
                        row = dt.NewRow();  //assign new row to DataTable
                        rowCounter = 0;
                        for (int j = 1; j <= colCount; j++) //Loop for available column of excel data
                        {
                            //check if cell is empty
                            //if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                          
                            dt.Rows.Add(rowCount);

                            // if (excelRange.Cells[i, j] != null)
                            //  {
                          //  dt.Rows.Add(row) = ((Excel.Range)excelRange.Cells[i, j]).Text.ToString();
                              //  row[i] = ((Excel.Range)excelRange.Cells[i, j]).Text.ToString();
                                //MessageBox.Show(((Excel.Range)excelRange.Cells[i, j]).Text.ToString());

                                //row[rowCounter] = excelRange.Cells[i, j].Value2.ToString();
                                // row[rowCounter] = excelRange.Value2 == null ? "" : excelRange.Value2.ToString();

                          //  }
                          //  else
                          //  {
                          //      row[i] = "";
                          //  }

                            MessageBox.Show(rowCounter++.ToString());
                            rowCounter++;
                        }
                        
                        dt.Rows.Add(row); //add row to DataTable
                    }

                    dataGridView1.DataSource = dt; //assign DataTable as Datasource for DataGridview

                    //Close and Clean excel process
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    Marshal.ReleaseComObject(excelRange);
                    Marshal.ReleaseComObject(excelWorksheet);
                    excelWorkbook.Close();
                    Marshal.ReleaseComObject(excelWorkbook);

                    //quit 
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string file = ""; //variable for the Excel File Location
            System.Data.DataTable dt = new System.Data.DataTable();   //container for our excel data
            DataRow row;
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Check if Result == "OK".
            {
                file = openFileDialog1.FileName; //get the filename with the location of the file
                try
                {
                    //Create Object for Microsoft.Office.Interop.Excel that will be use to read excel file

                    Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

                    Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(file);

                    Microsoft.Office.Interop.Excel._Worksheet excelWorksheet = excelWorkbook.Sheets[1];

                    Microsoft.Office.Interop.Excel.Range excelRange = excelWorksheet.UsedRange;

                    int rowCount = excelRange.Rows.Count; //get row count of excel data

                    int colCount = excelRange.Columns.Count; // get column count of excel data

                    //Get the first Column of excel file which is the Column Name

                    for (int i = 1; i <= rowCount; i++)
                    {
                        for (int j = 1; j <= colCount; j++)
                        {
                            // dt.Columns.Add(excelRange.Cells[i, j].Value2.ToString());
                            //Value2 數值為動態, 會一直不見, 所以改直接用Text取值
                            dt.Columns.Add(((Microsoft.Office.Interop.Excel.Range)excelRange.Cells[i, j]).Text.ToString());  

                        }
                        break;
                    }

                    //Get Row Data of Excel

                    int rowCounter; //This variable is used for row index number
                    for (int i = 2; i <= rowCount; i++) //Loop for available row of excel data
                    {
                        row = dt.NewRow(); //assign new row to DataTable
                        rowCounter = 0;
                        for (int j = 1; j <= colCount; j++) //Loop for available column of excel data
                        {
                            //check if cell is empty
                            //if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                            //{
                            //    row[rowCounter] = excelRange.Cells[i, j].Value2.ToString();
                            //}
                            //else
                            //{
                            //    row[i] = "";
                            //}
                            row[rowCounter] = ((Microsoft.Office.Interop.Excel.Range)excelRange.Cells[i, j]).Text.ToString();
                            rowCounter++;
                        }
                        dt.Rows.Add(row); //add row to DataTable
                    }

                    dataGridView1.DataSource = dt; //assign DataTable as Datasource for DataGridview

                    //close and clean excel process
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    Marshal.ReleaseComObject(excelRange);
                    Marshal.ReleaseComObject(excelWorksheet);
                    //quit apps
                    excelWorkbook.Close();
                    Marshal.ReleaseComObject(excelWorkbook);
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            }
    }
}
