using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_Sheet_Collator
{
    public partial class Form1 : Form
    {

        // Set default directory
        string currentDirectory = "C:\\";

        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Open the selected worksheet within the relevant excel file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listBoxFiles_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBoxFiles.SelectedItems.Count == 1)
            {
                // Instantiate Excel application
                Excel.Application xl2 = new Excel.Application();
                xl2.Visible = true; // Excel visible to user

                // Separate sheet name and file name for processing
                string fullName = listBoxFiles.GetItemText(listBoxFiles.SelectedItem);
                string[] stringSeparators = new string[] { "  -  " };
                string[] fileSheet = fullName.Split(stringSeparators, StringSplitOptions.None);
                string sheetName = fileSheet[0];
                string fileName = fileSheet[1];

                // Open an existing workbook
                string workbookPath = currentDirectory + fileName;
                Excel.Workbook wb2 = xl2.Workbooks.Open(workbookPath,
                    0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);

                // Get sheet in workbook
                Excel.Sheets excelSheets = wb2.Worksheets;
                Excel.Worksheet wsh = (Excel.Worksheet)excelSheets[sheetName];
                wsh.Activate(); // Activate the worksheet

                listBoxFiles.SelectedItems.Clear(); // Clear selected item
            }
        }

        /// <summary>
        /// Load/Refresh the Item List Asynchronously
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void buttonRefresh_Click(object sender, EventArgs e)
        {
            listBoxFiles.Items.Clear(); // clear all items after first refresh

            try
            {
                var progress = new Progress<string>((x) => listBoxFiles.Items.Add(x));
                await Task.Run(() => DoWork(progress));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        /// <summary>
        /// Background thread to prevent freezing IO with large number of files
        /// </summary>
        /// <param name="progress"></param>
        private void DoWork(IProgress<string> progress)
        {
            // Obtain files in directory
            DirectoryInfo dinfo = new DirectoryInfo(currentDirectory);
            FileInfo[] Files = dinfo.GetFiles("*.xlsx");

            // Instantiate Excel application
            Excel.Application xl = new Excel.Application();

            // Read all files
            foreach (FileInfo file in Files)
            {
                //listBoxFiles.Items.Add(file.Name); // uncomment to see all files in another listBox (need to another listBox)
                Excel.Workbook wb = xl.Workbooks.Open(currentDirectory + file.Name, 0, true);

                // Read and assign all sheets and relevant files
                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    string wsName = ws.Name;
                    string fileName = file.Name;
                    progress.Report(wsName + "  -  " + fileName);  
                }
            }
        }

        /// <summary>
        /// Search for specific file names or sheet names
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonSearch_Click(object sender, EventArgs e)
        {
            List<string> newItems = new List<string>();
            foreach (object item in listBoxFiles.Items)
            {
                // Check if search keyword matches with any items in the list
                if (item.ToString().ToLower().Contains(textBoxSearch.Text.ToLower())) // Case-insensitive
                {
                    newItems.Add(item.ToString());
                }
            }

            listBoxNewFiles.Items.Clear(); //Clear all items after first load

            // Add all matching items in the new ListBox
            foreach (string item in newItems)
            {
                listBoxNewFiles.Items.Add(item);
            }
        }

        /// <summary>
        /// /// Open the selected worksheet within the relevant excel file in the refined list
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listBoxNewFiles_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBoxNewFiles.SelectedItems.Count == 1)
            {
                // Instantiate Excel application
                Excel.Application xl2 = new Excel.Application();
                xl2.Visible = true; // Excel visible to user

                // Separate sheet name and file name for processing
                string fullName = listBoxNewFiles.GetItemText(listBoxNewFiles.SelectedItem);
                string[] stringSeparators = new string[] { "  -  " };
                string[] fileSheet = fullName.Split(stringSeparators, StringSplitOptions.None);
                string sheetName = fileSheet[0];
                string fileName = fileSheet[1];

                // Open an existing workbook
                string workbookPath = currentDirectory + fileName;
                Excel.Workbook wb2 = xl2.Workbooks.Open(workbookPath,
                    0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);

                // Get sheet in workbook
                Excel.Sheets excelSheets = wb2.Worksheets;
                Excel.Worksheet wsh = (Excel.Worksheet)excelSheets[sheetName];
                wsh.Activate(); // Activate the worksheet
                
                listBoxNewFiles.SelectedItems.Clear(); // Clear selected item
            }
        }

        /// <summary>
        /// Select the folder all excel files and sheets will be drawn from
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonSelectFolder_Click(object sender, EventArgs e)
        {
            // Instantiate folder browser dialog
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            // Update directory based on chosen folder
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                currentDirectory = fbd.SelectedPath + "\\";
        }
    }
}
