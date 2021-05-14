using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using ClosedXML;
using Excel = Microsoft.Office.Interop.Excel;

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
                IXLWorksheet ws = wb.Worksheets.Add();//In workbook create worksheet and give name
                                                      //To delete worksheet ---> wb.Worksheet("Sample Sheet").Delete();



               


                //Give headers using textbox inputs

                ws.Cell(1, 1).Value = box_0101.Text;
                ws.Cell(1, 2).Value = box_0102.Text;
                ws.Cell(1, 3).Value = box_0103.Text;
                ws.Cell(1, 4).Value = box_0104.Text;
                ws.Cell(1, 5).Value = box_0105.Text;
                ws.Cell(1, 6).Value = box_0106.Text;
                ws.Cell(1, 7).Value = box_0107.Text;
                ws.Cell(1, 8).Value = box_0108.Text;



                // Add a bunch of numbers to filter
                ws.Cell("a2").SetValue(10)
                             .CellBelow().SetValue(2)
                             .CellBelow().SetValue(3)
                             .CellBelow().SetValue(3)
                             .CellBelow().SetValue(5)
                             .CellBelow().SetValue(1)
                             .CellBelow().SetValue(4);
                ws.Cell("b2").SetValue("a1")
                             .CellBelow().SetValue("a")
                             .CellBelow().SetValue("b")
                             .CellBelow().SetValue("c")
                             .CellBelow().SetValue("d")
                             .CellBelow().SetValue("e")
                             .CellBelow().SetValue("f");

                // Add filters
                ws.RangeUsed().SetAutoFilter().Column(1).LessThan(4);
                ws.AutoFilter.Sort(1);



                wb.SaveAs(@"Y:\Liverpool projects\Windows form app\Ruth\test.xlsx");

            }

            
        }

 
    }
}




           