using System;
using System.IO;
using System.Data;
using EasyXMLSerializer;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using GemBox.Spreadsheet;
using ClosedXML;
using Excel = Microsoft.Office.Interop.Excel;
using Aspose.Cells;

namespace UsingClosedXML
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        
        private void button1_Click(object sender, EventArgs e)
        {
            //SELECT FILE AND STORE
            var fileContent = string.Empty;
            string[] filePath;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "y:\\";
                openFileDialog.Filter = "excel|*.xls; *.xlsx; ";
                openFileDialog.Multiselect = true;
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file, View it
                    filePath = openFileDialog.FileNames;
                    string filePathString = string.Join("\n", filePath);
                    Lbl_1.Text = filePathString;
                    //LOOP THROUGH FILES- if we choose to do this we'll probably want to use a separate output file for each file
                }

            }
        }

      
        private void button2_Click(object sender, EventArgs e)
        {

            string[] filePath = Lbl_1.Text.Split('\n');
            foreach (string sFileName in filePath)
            {
                //Create workbook used closedXML
                IXLWorkbook wb = new XLWorkbook(@"Y:\Liverpool projects\Windows form app\Ruth\Auto_OH166_Chip_1_Oncomine_Myeloid_hods-department_500.bcmatrix.xlsx");//Import file change to "sFileName"
                var ws = wb.Worksheet(1);                                                                                                                                                     //IXLWorksheet ws2 = wb.Worksheets.Add("Sample Sheet");//In workbook create worksheet and give name
                                                                                                                                                                                              //To delete worksheet ---> wb.Worksheet("Sample Sheet").Delete();


                // Add filters
                ws.RangeUsed().SetAutoFilter().Column(3).LessThan(100);
                ws.AutoFilter.Sort(3);


                //Copy files from import into new sheet

                var firstTableCell = ws.FirstCellUsed();
                var lastTableCell = ws.LastCellUsed();
                var rngData = ws.Range(firstTableCell.Address, lastTableCell.Address);
                var firstColumn = ws.FirstColumn();
                int numRows = rngData.RowCount();
                int numColumn = rngData.ColumnCount();

                // Create a file stream containing the Excel file to be opened
                FileStream fstream = new FileStream(@"Y:\Liverpool projects\Windows form app\Ruth\Auto_OH166_Chip_1_Oncomine_Myeloid_hods-department_500.bcmatrix.xlsx", FileMode.Open);

                // Instantiate a Workbook object
                //Opening the Excel file through the file stream
                Workbook workbook = new Workbook(fstream);

                // Access the first worksheet in the Excel file
                Worksheet worksheet = workbook.Worksheets[0];

                // Export the contents of 2 rows and 2 columns starting from 1st cell to DataTable
                DataTable dataTable = worksheet.Cells.ExportDataTable(0, 0, numRows, numColumn, true);

                //Filter
                var results = dataTable.Select(@"Ioncode_0101 < 100");
                
                var results2 = dataTable.Select(@"Ioncode_0102 < 100");
                var results3 = dataTable.Select(@"Ioncode_0103 < 100");
                var results4 = dataTable.Select(@"Ioncode_0104 < 100");
                var results5 = dataTable.Select(@"Ioncode_0105 < 100");
                var results6 = dataTable.Select(@"Ioncode_0106 < 100");
                var results7 = dataTable.Select(@"Ioncode_0107 < 100");
                var results8 = dataTable.Select(@"Ioncode_0108 < 100");



                // Bind the DataTable with DataGrid
                dataGridView1.DataSource = dataTable;

                // Close the file stream to free all resources
                fstream.Close();



                




                // Copy the table to another worksheet
                var wsCopy = wb.Worksheets.Add("ws2");
                wsCopy.Cell(1, 1).Value = rngData;

                var wsCopy2 = wb.Worksheets.Add("ws3");
                wsCopy2.Cell(1, 1).Value = firstColumn;

                var wsCopy3 = wb.Worksheets.Add("ws4");
                wsCopy3.Cell(1, 1).Value = results;

                //Copying first row in IonCode_0101 and pasting it to another ws
                var geneResult1 = wsCopy3.RangeUsed().Column(1);
                var wsCopy4 = wb.Worksheets.Add("ws5");
                wsCopy4.Cell(1, 1).Value = geneResult1;





                //ws.Cell(1, 1).Value = box_0101.Text;
                //ws.Cell(1, 2).Value = box_0102.Text;
                //ws.Cell(1, 3).Value = box_0103.Text;
                //ws.Cell(1, 4).Value = box_0104.Text;
                //ws.Cell(1, 5).Value = box_0105.Text;
                //ws.Cell(1, 6).Value = box_0106.Text;
                //ws.Cell(1, 7).Value = box_0107.Text;
                //ws.Cell(1, 8).Value = box_0108.Text;


                wb.SaveAs(@"Y:\Liverpool projects\Windows form app\Ruth\test.xlsx");

            }

            
        }

 
    }
}




           