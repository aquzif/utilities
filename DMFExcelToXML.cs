using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UTILS
{
    public static class DMFExcelToXML
    {
        public static async Task Run(Form parentForm)
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
            parentForm.Controls.Add(progressBar);

            await Task.Run(() => ((Form1)parentForm).ProcessExcelToXml(entityName, path, saveDirectory, baseFileName, progressBar, partXml == DialogResult.Yes));
        }
    }
}
