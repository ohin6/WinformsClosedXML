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
                var Ioncode_0101 = dataTable.Select(@"Ioncode_0101 < 100");
                var Ioncode_0102 = dataTable.Select(@"Ioncode_0102 < 100");
                var Ioncode_0103 = dataTable.Select(@"Ioncode_0103 < 100");
                var Ioncode_0104 = dataTable.Select(@"Ioncode_0104 < 100");
                var Ioncode_0105 = dataTable.Select(@"Ioncode_0105 < 100");
                var Ioncode_0106 = dataTable.Select(@"Ioncode_0106 < 100");
                var Ioncode_0107 = dataTable.Select(@"Ioncode_0107 < 100");
                var Ioncode_0108 = dataTable.Select(@"Ioncode_0108 < 100");

                // Close the file stream to free all resources
                fstream.Close();



                // Copy the table to another worksheet THIS SECTION IS NOT RELEVANT
                var wsCopy = wb.Worksheets.Add("ws2");
                wsCopy.Cell(1, 1).Value = rngData;

                var wsCopy2 = wb.Worksheets.Add("ws3");
                wsCopy2.Cell(1, 1).Value = firstColumn;

                
                //
                //Paste datatable queries to Worksheet (ws4) and paste gene target into worksheet (ws5)
                
                //1. create worksheet ws4 and paste dataTable query result (line 95 to 102)
                // first query is Ioncode_0101
                var wsCopy3 = wb.Worksheets.Add("ws4");
                wsCopy3.Cell(1, 1).Value = Ioncode_0101; //paste row1 column 1

                //2. copy gene target (column 1) and paste into new worksheet (ws5)
                var geneIoncode = wsCopy3.RangeUsed().Column(1); 
                var wsCopy4 = wb.Worksheets.Add("ws5");
                wsCopy4.Cell(1,1).Value = geneIoncode;

                //3. Repeat this process for all database queries - should be made into a loop
                //3a. Ioncode_0102
                wb.Worksheet("ws4").Clear();     // clear previous query results from ws4           
                wsCopy3.Cell(1, 1).Value = Ioncode_0102; //paste new query to ws4
                wsCopy4.Cell(1, 2).Value = geneIoncode; //paste gene target from ws4 to ws5 column 2

                //3b. Ioncode_0103
                wb.Worksheet("ws4").Clear();
                wsCopy3.Cell(1, 1).Value = Ioncode_0103;
                wsCopy4.Cell(1, 3).Value = geneIoncode;

                //3c. Ioncode_0104
                wb.Worksheet("ws4").Clear();
                wsCopy3.Cell(1, 1).Value = Ioncode_0104;
                wsCopy4.Cell(1, 4).Value = geneIoncode;

                //3d. Ioncode_0105 
                wb.Worksheet("ws4").Clear();
                wsCopy3.Cell(1, 1).Value = Ioncode_0105;
                wsCopy4.Cell(1, 5).Value = geneIoncode;

                //3e. Ioncode_0106 
                wb.Worksheet("ws4").Clear();
                wsCopy3.Cell(1, 1).Value = Ioncode_0106;
                wsCopy4.Cell(1, 6).Value = geneIoncode;

                //3f. Ioncode_0107 
                wb.Worksheet("ws4").Clear();
                wsCopy3.Cell(1, 1).Value = Ioncode_0107;
                wsCopy4.Cell(1, 7).Value = geneIoncode;

                //3g. Ioncode_0108
                wb.Worksheet("ws4").Clear();
                wsCopy3.Cell(1, 1).Value = Ioncode_0108;
                wsCopy4.Cell(1, 8).Value = geneIoncode;



                //Insert sample names as Headers from text file
                var Worksheet5 = wb.Worksheet(5); //making ws5 a variable called Worksheet5 as this allows to modify
                Worksheet5.Cell(1, 1).Value = box_0101.Text;
                Worksheet5.Cell(1, 2).Value = box_0102.Text;
                Worksheet5.Cell(1, 3).Value = box_0103.Text;
                Worksheet5.Cell(1, 4).Value = box_0104.Text;
                Worksheet5.Cell(1, 5).Value = box_0105.Text;
                Worksheet5.Cell(1, 6).Value = box_0106.Text;
                Worksheet5.Cell(1, 7).Value = box_0107.Text;
                Worksheet5.Cell(1, 8).Value = box_0108.Text;

                //Transpose data
                var Worksheet5Range = Worksheet5.RangeUsed();
                Worksheet5Range.Transpose(XLTransposeOptions.MoveCells);
                Worksheet5.Columns().AdjustToContents();


                wb.SaveAs(@"Y:\Liverpool projects\Windows form app\Ruth\test.xlsx");

            }

            
        }

 
    }
}




           