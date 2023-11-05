using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.PowerPoint;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Ppt = Microsoft.Office.Interop.PowerPoint;
using static System.Net.WebRequestMethods;
using Aspose.Slides.Export.Web;
using Document = iTextSharp.text.Document;
using Image = iTextSharp.text.Image;
using File = System.IO.File;

namespace ToPDF
{
    public partial class frmConvert : Form
    {
        private List<string> fpath = new List<string>();
        public Boolean ConvertStatus = true;
        public Boolean imgToPDFStatus;

        public frmConvert()
        {
            InitializeComponent();
        }

        private void btnChooseFiles_Click(object sender, EventArgs e)
        {
            //bool isEmpty = !fpath.Any();

            //include openFileDialog from toolbox
            openFileDialog1.ShowDialog(this);
            foreach (string file in openFileDialog1.FileNames)
            {
                lstDoc.Items.Add(file);
                fpath.Add(file);

                if (Path.GetExtension(file).ToLower() == ".jpg" ||
                    Path.GetExtension(file).ToLower() == ".jpeg" ||
                    Path.GetExtension(file).ToLower() == ".png" ||
                    Path.GetExtension(file).ToLower() == ".bmp")
                {
                    txtFilename.Enabled = true;
                    txtFilename.Text = null;
                }


            }

            //Display path for from
            lblfrom.Text = openFileDialog1.FileName;
            


            //if no file selected, convert button is still disable
            if (fpath.Count != 0)
            {
                btnSelectFolder.Enabled = true;
            } 
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            openFileDialog1.FileName = null;
            lstDoc.Items.Clear();
            fpath.Clear();
            btnConvert.Enabled = false;
            btnSelectFolder.Enabled=false;
            txtFilename.Enabled = false;
            lblfrom.Text = null;
            lblTo.Text = null;
            txtFilename.Text = "(For Image to PDF file name is required)";
        }

        private void WordToPDF(string file, Word.Application wordApp, string savePath)
        {
            string OutputPath, InputPath;
            InputPath = Path.ChangeExtension(file, ".docx");
            Word.Document wordDoc = wordApp.Documents.Open(InputPath);

            //change extension of the file to .pdf before saving it
            OutputPath = Path.ChangeExtension(file, ".pdf");
            OutputPath = savePath + Path.GetFileName(OutputPath);
            if (checkExistFile(OutputPath))
            {
                string message = Path.GetFileName(OutputPath) + " already exists.\n" + "Do you want to replace it ?";
                string title = "Confirm Save As";
                DialogResult dialogResult = MessageBox.Show(message, title, MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if(dialogResult == DialogResult.Yes)
                {
                    wordDoc.SaveAs2(OutputPath, Word.WdSaveFormat.wdFormatPDF);
                }
                else
                {
                    ConvertStatus = false;
                }
            }
            else
            {
                wordDoc.SaveAs2(OutputPath, Word.WdSaveFormat.wdFormatPDF);
            }
            wordDoc.Close();
        }

        private void PptToPDF(string file, Ppt.Application pptApp, string savePath)
        {
            string OutputPath;
            try
            {
                //InputPath = Path.ChangeExtension(file, ".pptx");
                OutputPath = Path.ChangeExtension(file, ".pdf");
                OutputPath = savePath + Path.GetFileName(OutputPath);
                var pptPresentation = pptApp.Presentations.Open(file, WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse);
                if (checkExistFile(OutputPath))
                {
                    string message = Path.GetFileName(OutputPath) + " already exists.\n" + "Do you want to replace it ?";
                    string title = "Confirm Save As";
                    DialogResult dialogResult = MessageBox.Show(message, title, MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (dialogResult == DialogResult.Yes)
                    {
                        pptPresentation.ExportAsFixedFormat(OutputPath, PpFixedFormatType.ppFixedFormatTypePDF);
                    }
                    else
                    {
                        ConvertStatus = false;
                    }
                }
                else
                {
                    pptPresentation.ExportAsFixedFormat(OutputPath, PpFixedFormatType.ppFixedFormatTypePDF);
                }
                

                pptPresentation.Close();
            }
            catch
            {
                ConvertStatus = false;
                // Handle any exceptions that might occur
                MessageBox.Show("The following Powerpoint file cannot be converted:\n"
                    +file, "Empty File", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExcelToPDF(string file, Excel.Application excelApp, string savePath)
        {
            string OutputPath, InputPath;

            InputPath = Path.ChangeExtension(file, ".xlsx");
            OutputPath = Path.ChangeExtension(file, ".pdf");
            OutputPath = savePath + Path.GetFileName(OutputPath);
            // Open Excel file
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(InputPath);

            if (checkExistFile(OutputPath))
            {
                string message = Path.GetFileName(OutputPath) + " already exists.\n" + "Do you want to replace it ?";
                string title = "Confirm Save As";
                DialogResult dialogResult = MessageBox.Show(message, title, MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dialogResult == DialogResult.Yes)
                {
                    // Save Excel file as PDF
                    excelWorkbook.ExportAsFixedFormat(
                        Type: Excel.XlFixedFormatType.xlTypePDF,
                        Filename: OutputPath,
                        Quality: Excel.XlFixedFormatQuality.xlQualityStandard,
                        IncludeDocProperties: true,
                        IgnorePrintAreas: false,
                        OpenAfterPublish: false);
                }
                else
                {
                    ConvertStatus = false;
                }
            }
            else
            {
                // Save Excel file as PDF
                excelWorkbook.ExportAsFixedFormat(
                    Type: Excel.XlFixedFormatType.xlTypePDF,
                    Filename: OutputPath,
                    Quality: Excel.XlFixedFormatQuality.xlQualityStandard,
                    IncludeDocProperties: true,
                    IgnorePrintAreas: false,
                    OpenAfterPublish: false);
            }
            

            // Close Excel file and quit Excel application
            excelWorkbook.Close();
        }

        private void ImageToPDF(string file, Document document, string savePath)
        {
            
            Image image = Image.GetInstance(file);
            float width = image.Width;
            float height = image.Height;
            if (width > height && width > document.PageSize.Width)
            {
                height *= document.PageSize.Width / width;
                width = document.PageSize.Width;
            }
            else if (height > width && height > document.PageSize.Height)
            {
                width *= document.PageSize.Height / height;
                height = document.PageSize.Height;
            }
            else if (width == height && width > document.PageSize.Width && height > document.PageSize.Height)
            {
                height *= document.PageSize.Width / width;
                width = document.PageSize.Width;
            }

            image.ScaleAbsolute(width, height);
            document.Add(image);

        }

        private Boolean checkExistFile(string OutputPath)
        {
            if (File.Exists(OutputPath))
            {
                return true;
            }
            return false;
        }

        
        private void btnConvert_Click(object sender, EventArgs e)
        {
            string savePath = lblTo.Text;
            
            //install Microsoft.Office.Interop.Word in the Mangage Nuget packages before using the library
            //using Word = Microsoft.Office.Interop.Word; (look above)
            //declare this here to make program run faster 
            Word.Application wordApp = new Word.Application();
            Ppt.Application pptApp = new Ppt.Application();
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;
            Document document = new Document();

            string filename= txtFilename.Text;
            

            //using iTextSharp Document
            //Document document;

            if (!string.IsNullOrEmpty(filename))
            {
                if (txtFilename.Enabled == true)
                {
                    // + DateTime.Now.ToString("M-d-yyyy"
                    imgToPDFStatus = true;
                    string outputFilePath = Path.Combine(savePath, Path.ChangeExtension(filename, ".pdf"));
                    if (checkExistFile(outputFilePath))
                    {
                        string message = Path.GetFileName(outputFilePath) + " already exists\n" + "Do you want to replace it ?";
                        string title = "Confirm Save As";
                        DialogResult dialogResult = MessageBox.Show(message, title, MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        if (dialogResult == DialogResult.Yes)
                        {
                            PdfWriter.GetInstance(document, new FileStream(outputFilePath, FileMode.Create));
                            document.Open();
                        }
                        else
                        {
                            ConvertStatus = false;
                            imgToPDFStatus = false;
                        }
                    }
                    else
                    {
                        PdfWriter.GetInstance(document, new FileStream(outputFilePath, FileMode.Create));
                        document.Open();
                    }
                    
                }
                foreach (string file in fpath)
                {
                    if (Path.GetExtension(file) == ".docx")
                    {
                        WordToPDF(file, wordApp, savePath);
                    }
                    else if (Path.GetExtension(file) == ".pptx"|| Path.GetExtension(file) == ".ppt")
                    {
                        PptToPDF(file, pptApp, savePath);
                    }
                    else if (Path.GetExtension(file).ToLower() == ".jpg" ||
                        Path.GetExtension(file).ToLower() == ".jpeg" ||
                        Path.GetExtension(file).ToLower() == ".png" ||
                        Path.GetExtension(file).ToLower() == ".bmp")
                    {
                        if(imgToPDFStatus==true)
                            ImageToPDF(file, document, savePath);
                    }
                    else if (Path.GetExtension(file) == ".xlsx")
                    {
                        ExcelToPDF(file, excelApp, savePath);
                    }

                }
               
                document.Close();
                wordApp.Quit();
                excelApp.Quit();
                pptApp.Quit();

                if (ConvertStatus == true)
                {
                    MessageBox.Show("Converted Successfully !");
                }

                //set openFiledialog filename to null to avoid the remaining names of previous files
                btnClear_Click(sender, e);
            }
            else
            {
                MessageBox.Show("File name is required!");
            }
        }

        private void btnSelectFolder_Click(object sender, EventArgs e)
        {
            ConvertStatus = true;
            if (folderBrowserDialog1.ShowDialog() != DialogResult.Cancel)
            {
                lblTo.Text = folderBrowserDialog1.SelectedPath.ToString() + "\\";
                btnConvert.Enabled = true;
                
            }
       
        }

        //Before using DragDrop you must set allowdrop to true in the listbox property setting!
        private void lstDoc_DragDrop(object sender, DragEventArgs e)
        {
            lstDoc.BackColor = Color.Snow;
            
            //Allows to drop a file into the listbox then display each path of the file
            string[] filePath = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            
            //display file path for from
            lblfrom.Text = filePath[0];

            foreach (string file in filePath)
            {
                if (Path.GetExtension(file) == ".docx" || Path.GetExtension(file) == ".pptx" || Path.GetExtension(file) == ".ppt" || Path.GetExtension(file) == ".xlsx")
                {
                    lstDoc.Items.Add(file);
                    fpath.Add(file);
                    btnSelectFolder.Enabled = true;

                }else if(Path.GetExtension(file).ToLower() == ".jpg" ||
                    Path.GetExtension(file).ToLower() == ".jpeg" ||
                    Path.GetExtension(file).ToLower() == ".png" ||
                    Path.GetExtension(file).ToLower() == ".bmp")
                {
                    lstDoc.Items.Add(file);
                    fpath.Add(file);
                    txtFilename.Enabled = true;
                    txtFilename.Text = null;
                    btnSelectFolder.Enabled = true;
                }
                else
                {
                    MessageBox.Show("Cannot Convert " + Path.GetFileName(file) + "\n"
                        + "Please select the files correctly!", "Wrong Format",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void lstDoc_DragEnter(object sender, DragEventArgs e)
        {
            //Drag and drop effects in windows
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false) == true)
                e.Effect = DragDropEffects.All;    
        }

        private void lstDoc_DragOver(object sender, DragEventArgs e)
        {
            lstDoc.BackColor = Color.LightGray;
        }

        private void lstDoc_DragLeave(object sender, EventArgs e)
        {
            lstDoc.BackColor = Color.Snow;
        }

        private void btnRemove_Click(object sender, EventArgs e)
        {

            if (lstDoc.SelectedIndex > -1)
            {
                int rmv = lstDoc.SelectedIndex;
                lstDoc.Items.RemoveAt(rmv);
                fpath.RemoveAt(rmv);

                if (imageExist() == false)
                {
                    txtFilename.Enabled = false;
                    txtFilename.Text = "(For Image to PDF file name is required)";
                }

                if (lstDoc.Items.Count == 0)
                {
                    btnClear_Click(sender, e);
                }
                
            }
            else
            {
                MessageBox.Show("Select an item to remove!");
            }
        }

        private Boolean imageExist()
        {
            foreach (string file in fpath)
            {
                if(Path.GetExtension(file).ToLower() == ".jpg" ||
                    Path.GetExtension(file).ToLower() == ".jpeg" ||
                    Path.GetExtension(file).ToLower() == ".png" ||
                    Path.GetExtension(file).ToLower() == ".bmp")
                {
                    return true;
                }
            }
            return false;
        }

        private void lstDoc_KeyDown(object sender, KeyEventArgs e)
        {
            
            if(e.KeyCode == Keys.Delete)
            {
                btnRemove_Click(sender, e);
            }
            
        }
    }
}
