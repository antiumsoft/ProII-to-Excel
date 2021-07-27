using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProII_to_Excel
{
    public partial class OpenFileForm : Form
    {
        public OpenFileForm()
        {
            InitializeComponent();
        }

        private void loadButton_Click(object sender, EventArgs e)
        {
            // Create an instance of the open file dialog box.
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            // Set filter options and filter index.
            openFileDialog1.Filter = "Pro II PLT Files|*.plt";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.Multiselect = false;

            // Call the ShowDialog method to show the dialog box.
            DialogResult userClickedOK = openFileDialog1.ShowDialog();

            // Process input if the user clicked OK.
            if (userClickedOK == DialogResult.OK)
            {
                // Open the selected file to read.
                System.IO.Stream fileStream = openFileDialog1.OpenFile();
                IWorkbook wb = new Conversion(fileStream).Convert();
                if (wb != null) saveFile(wb);
            }
        }

        private void saveFile(IWorkbook wb)
        {

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "EXCEL 2007+ File |*.xlsx";
            saveFileDialog1.Title = "Save the generated Excel File";
            saveFileDialog1.ShowDialog();
            // If the file name is not an empty string open it for saving.
            if (saveFileDialog1.FileName != "")
            {
                // Saves the file via a FileStream created by the OpenFile method.
                System.IO.FileStream fs =
                           (System.IO.FileStream)saveFileDialog1.OpenFile();
                wb.Write(fs);
                try
                {
                    fs.Close();
                }
                catch (Exception ex) { }
                MessageBox.Show("File has been saved", "User Support");
            }
        }

        private void infoBox_Click(object sender, EventArgs e)
        {
            AboutBox aboutBox = new AboutBox();
            aboutBox.Show();
        }

    }
}
