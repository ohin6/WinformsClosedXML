using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using ClosedXML.Excel;
using Aspose.Cells;

namespace UsingClosedXML
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            /// Initialise lists of all text boxes and respective labels
            List<System.Windows.Forms.TextBox> textBoxes = new List<System.Windows.Forms.TextBox>
            {
                box_0101, box_0102, box_0103, box_0104, box_0105, box_0106, box_0107, box_0108
            };
            List<System.Windows.Forms.Label> labels = new List<System.Windows.Forms.Label>
            {
                label_1, label_2, label_3, label_4, label_5, label_6, label_7, label_8
            };

            /// hide the boxes and labels before the user selected number of samples
            for (int i = 0; i < 8; i++)
            {
                textBoxes[i].Hide();
                labels[i].Hide();
            }

            //Hide other parameters
            Lbl_samples.Hide();
            comboBox1.Hide();
            button2.Hide();
            warningLabel.Hide();

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

            //unhide 
            Lbl_samples.Show();
            comboBox1.Show();
            button2.Show();
        }

      
        private void button2_Click(object sender, EventArgs e)
        {

            string[] filePath = Lbl_1.Text.Split('\n');
            foreach (string sFileName in filePath)
            {
                //Create workbook used closedXML
                IXLWorkbook wb = new XLWorkbook(sFileName);//Import file change to "sFileName"
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
                FileStream fstream = new FileStream(sFileName, FileMode.Open);

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

            //Show message saved
            warningLabel.Show();
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<System.Windows.Forms.TextBox> textBoxes = new List<System.Windows.Forms.TextBox>
            {
                box_0101, box_0102, box_0103, box_0104, box_0105, box_0106, box_0107, box_0108
            };

            List<System.Windows.Forms.Label> labels = new List<System.Windows.Forms.Label>
            {
                label_1, label_2, label_3, label_4, label_5, label_6, label_7, label_8
            };

            /// hide the boxes and labels before the user selected number of samples
            for (int i = 0; i < 8; i++)
            {
                textBoxes[i].Hide();
                labels[i].Hide();
            }

            for (int i = 0; i < comboBox1.SelectedIndex + 1; i++)
            {
                textBoxes[i].Show();
                labels[i].Show();

            }
        }

        private void box_0101_TextChanged(object sender, EventArgs e)
        {
            if (box_0101.Text == "")
            {
                return;
            }
            /// if text is numerical and length = 5 then the input is valid
            if (int.TryParse(box_0101.Text, out _) && box_0101.Text.Length == 5)
            {
                warningLabel_1.ForeColor = Color.Green;
                warningLabel_1.Text = "Valid Input";
            }
            else
            {
                warningLabel_1.ForeColor = Color.Red;
                warningLabel_1.Text = "Invalid Input";
            }

        }

        private void box_0102_TextChanged(object sender, EventArgs e)
        {
            if (box_0102.Text == "")
            {
                return;
            }
            /// if text is numerical and length = 5 then the input is valid
            if (int.TryParse(box_0102.Text, out _) && box_0102.Text.Length == 5)
            {
                warningLabel_2.ForeColor = Color.Green;
                warningLabel_2.Text = "Valid Input";
            }
            else
            {
                warningLabel_2.ForeColor = Color.Red;
                warningLabel_2.Text = "Invalid Input";
            }
        }

        private void box_0103_TextChanged(object sender, EventArgs e)
        {
            if (box_0103.Text == "")
            {
                return;
            }
            /// if text is numerical and length = 5 then the input is valid
            if (int.TryParse(box_0103.Text, out _) && box_0103.Text.Length == 5)
            {
                warningLabel_3.ForeColor = Color.Green;
                warningLabel_3.Text = "Valid Input";
            }
            else
            {
                warningLabel_3.ForeColor = Color.Red;
                warningLabel_3.Text = "Invalid Input";
            }
        }

        private void box_0104_TextChanged(object sender, EventArgs e)
        {
            if (box_0104.Text == "")
            {
                return;
            }
            /// if text is numerical and length = 5 then the input is valid
            if (int.TryParse(box_0104.Text, out _) && box_0104.Text.Length == 5)
            {
                warningLabel_4.ForeColor = Color.Green;
                warningLabel_4.Text = "Valid Input";
            }
            else
            {
                warningLabel_4.ForeColor = Color.Red;
                warningLabel_4.Text = "Invalid Input";
            }
        }

        private void box_0105_TextChanged(object sender, EventArgs e)
        {
            if (box_0105.Text == "")
            {
                return;
            }
            /// if text is numerical and length = 5 then the input is valid
            if (int.TryParse(box_0105.Text, out _) && box_0105.Text.Length == 5)
            {
                warningLabel_5.ForeColor = Color.Green;
                warningLabel_5.Text = "Valid Input";
            }
            else
            {
                warningLabel_5.ForeColor = Color.Red;
                warningLabel_5.Text = "Invalid Input";
            }
        }

        private void box_0106_TextChanged(object sender, EventArgs e)
        {
            if (box_0106.Text == "")
            {
                return;
            }
            /// if text is numerical and length = 5 then the input is valid
            if (int.TryParse(box_0106.Text, out _) && box_0106.Text.Length == 5)
            {
                warningLabel_6.ForeColor = Color.Green;
                warningLabel_6.Text = "Valid Input";
            }
            else
            {
                warningLabel_6.ForeColor = Color.Red;
                warningLabel_6.Text = "Invalid Input";
            }
        }

        private void box_0107_TextChanged(object sender, EventArgs e)
        {
            if (box_0107.Text == "")
            {
                return;
            }
            /// if text is numerical and length = 5 then the input is valid
            if (int.TryParse(box_0107.Text, out _) && box_0107.Text.Length == 5)
            {
                warningLabel_7.ForeColor = Color.Green;
                warningLabel_7.Text = "Valid Input";
            }
            else
            {
                warningLabel_7.ForeColor = Color.Red;
                warningLabel_7.Text = "Invalid Input";
            }
        }

        private void box_0108_TextChanged(object sender, EventArgs e)
        {
            if (box_0108.Text == "")
            {
                return;
            }
            /// if text is numerical and length = 5 then the input is valid
            if (int.TryParse(box_0108.Text, out _) && box_0108.Text.Length == 5)
            {
                warningLabel_8.ForeColor = Color.Green;
                warningLabel_8.Text = "Valid Input";
            }
            else
            {
                warningLabel_8.ForeColor = Color.Red;
                warningLabel_8.Text = "Invalid Input";
            }
        }
    }
}




           