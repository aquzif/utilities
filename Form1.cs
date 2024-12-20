using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace UTILS
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            // Ask for entity name
            string entityName = Microsoft.VisualBasic.Interaction.InputBox("Enter the entity name", "Entity Name", "INVENTPRODUCTDEFAULTORDERSETTINGSENTITY");

            // Input excel file
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog1.Title = "Select an Excel File";
            if (openFileDialog1.ShowDialog() != DialogResult.OK)
                return;

            string path = openFileDialog1.FileName;

            // Ask for the initial save file path
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "XML Files|*.xml";
            saveFileDialog1.Title = "Save the XML File";
            saveFileDialog1.FileName = entityName + ".xml";
            if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                return;

            string saveDirectory = Path.GetDirectoryName(saveFileDialog1.FileName);
            string baseFileName = Path.GetFileNameWithoutExtension(saveFileDialog1.FileName);

            // Ask if the XML should be parted
            DialogResult partXml = MessageBox.Show("Do you want to split the XML into parts of 1000 rows?", "Split XML", MessageBoxButtons.YesNo);

            // Initialize progress bar
            ProgressBar progressBar = new ProgressBar();
            progressBar.Minimum = 0;
            progressBar.Step = 1;
            progressBar.Dock = DockStyle.Bottom;
            this.Controls.Add(progressBar);

            await Task.Run(() => ProcessExcelToXml(entityName, path, saveDirectory, baseFileName, progressBar, partXml == DialogResult.Yes));
        }

        private void ProcessExcelToXml(string entityName, string path, string saveDirectory, string baseFileName, ProgressBar progressBar, bool splitXml)
        {
            Application xlApp = null;
            Workbook xlWorkbook = null;
            _Worksheet xlWorksheet = null;
            Range xlRange = null;

            try
            {
                // Open Excel file
                xlApp = new Application();
                xlWorkbook = xlApp.Workbooks.Open(path);
                xlWorksheet = xlWorkbook.Sheets[1];
                xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                // Ask which columns are numbers
                List<int> numericColumns = new List<int>();
                string numericColumnsInput = Microsoft.VisualBasic.Interaction.InputBox("Enter column numbers that contain numeric values, separated by commas", "Numeric Columns", "");
                if (!string.IsNullOrEmpty(numericColumnsInput))
                {
                    foreach (var col in numericColumnsInput.Split(','))
                    {
                        if (int.TryParse(col.Trim(), out int colIndex))
                        {
                            numericColumns.Add(colIndex);
                        }
                    }
                }

                // Update progress bar maximum value
                this.Invoke((MethodInvoker)delegate
                {
                    progressBar.Maximum = rowCount - 1;
                });

                int fileCounter = 1;
                int rowCounter = 0;
                string xml = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n<Document>";

                for (int i = 2; i <= rowCount; i++)
                {
                    bool isEmptyRow = true;
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (xlRange.Cells[i, j]?.Value2 != null)
                        {
                            isEmptyRow = false;
                            break;
                        }
                    }

                    if (isEmptyRow) continue;

                    xml += "<" + entityName + ">";
                    for (int j = 1; j <= colCount; j++)
                    {
                        string header = xlRange.Cells[1, j]?.Value2?.ToString();
                        string value = xlRange.Cells[i, j]?.Value2?.ToString();

                        if (numericColumns.Contains(j) && !string.IsNullOrEmpty(value))
                        {
                            value = value.Replace(",", ".");
                        }

                        if (!string.IsNullOrEmpty(header) && !string.IsNullOrEmpty(value))
                        {
                            xml += "<" + header + ">" + value + "</" + header + ">";
                        }
                    }
                    xml += "</" + entityName + ">";

                    rowCounter++;

                    // Save XML file in parts of 1000 if splitXml is true
                    if (splitXml && rowCounter >= 1000)
                    {
                        xml += "</Document>";
                        string partFileName = Path.Combine(saveDirectory, $"{baseFileName}.part{fileCounter}.xml");
                        File.WriteAllText(partFileName, xml);

                        fileCounter++;
                        rowCounter = 0;
                        xml = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n<Document>";
                    }

                    // Update progress bar
                    this.Invoke((MethodInvoker)delegate
                    {
                        progressBar.PerformStep();
                    });
                }

                // Save any remaining rows or the whole file if not splitting
                if (rowCounter > 0 || !splitXml)
                {
                    xml += "</Document>";
                    string partFileName = splitXml ? Path.Combine(saveDirectory, $"{baseFileName}.part{fileCounter}.xml") : Path.Combine(saveDirectory, $"{baseFileName}.xml");
                    File.WriteAllText(partFileName, xml);
                }
            }
            finally
            {
                // Release COM objects
                if (xlWorkbook != null)
                {
                    xlWorkbook.Close(false);
                    Marshal.ReleaseComObject(xlWorkbook);
                }
                if (xlWorksheet != null)
                    Marshal.ReleaseComObject(xlWorksheet);
                if (xlRange != null)
                    Marshal.ReleaseComObject(xlRange);
                if (xlApp != null)
                {
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);
                }

                // Hide progress bar when done
                this.Invoke((MethodInvoker)delegate
                {
                    this.Controls.Remove(progressBar);
                    MessageBox.Show("Processing complete!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
                });
            }
        }
    }
}
