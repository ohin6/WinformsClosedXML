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

            if (Lbl_1.Text == "Please Select a File")
            {
                Lbl_warning.ForeColor = Color.Red;
                Lbl_warning.Text = "Please select a file before continuing";
                return;
            }

            Lbl_warning.Text = "";

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "excel|*.xls";
                openFileDialog.Multiselect = true;
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file, View it
                    filePath = openFileDialog.FileNames;
                    string filePathString = string.Join("\n", filePath);
                    Lbl_1.Text = filePathString;
                }
                
            }
        }

      
        private void button2_Click(object sender, EventArgs e)
        {
            IXLWorkbook wb = new XLWorkbook();//Create workbook used closedXML
            IXLWorksheet ws = wb.Worksheets.Add("Sample Sheet");//In workbook create worksheet and give name
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

            //select range in cells
            //IXLRange rng = ws.Range("A1:H1");
            var firstCell = ws.FirstCellUsed();
            var lastCell = ws.LastCellUsed();
            var range = ws.Range(firstCell.Address, lastCell.Address);

            range.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;


            ws.Row(1).Sort();

            wb.SaveAs(@"Y:\Liverpool projects\Windows form app\Ruth\test.xlsx");
        }

 
    }
}




           