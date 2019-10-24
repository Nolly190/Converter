using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp.text;
using Bytescout.Spreadsheet;
using iTextSharp.text.pdf;

namespace Lonzec_Txt_Converter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private  void button1_Click(object sender, EventArgs e)
        {
            try
            {
                using (var fileDialog = new OpenFileDialog() { Filter = "Text File |*.txt", Multiselect = false })
                {
                    if (fileDialog.ShowDialog() == DialogResult.OK)
                    {
                       
                        lblFilePath.Text = fileDialog.FileName;
                      
                          var ext = System.IO.Path.GetExtension(fileDialog.FileName);
                          if (ext!=".txt")
                          {
                              MessageBox.Show("Select a txt file");
                              lblFilePath.Text = "";
                          }

                    }
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
           
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            rbExcel.Checked = true;
        }

        private async void btnConvert_Click(object sender, EventArgs e)
        {
            try
            {
               
                string readString = null;
                using (
                    var streamReader = new StreamReader(lblFilePath.Text))
                {
                    readString = await streamReader.ReadLineAsync();
                    var newList = new List<Task<string>>();
                    while (!streamReader.EndOfStream)
                    {
                        newList.Add(streamReader.ReadLineAsync());
                    }
                    var getExactUrl = System.IO.Path.GetDirectoryName(lblFilePath.Text);

                    if (rbPdf.Checked)
                    {
                        var size = newList[0].Result.Split('|');
                        Document pdfDocument = new Document();
                        PdfPTable pdfTable = new PdfPTable(size.Length);
                        int num = 1;
                        foreach (Task<string> item in newList)
                        {
                            size = item.Result.Split('|');

                            for (int i = 0; i < size.Length; i++)
                            {

                                pdfTable.AddCell(size[i] ?? " ");

                            }

                            num++;
                        }
                        if (File.Exists(getExactUrl + "ConvertedPdf.pdf"))
                        {
                            File.Delete(getExactUrl + "ConvertedPdf.pdf");
                        }
                        using (var fileStream = new FileStream(getExactUrl + "ConvertedPdf.pdf", FileMode.Create))
                        {


                            PdfWriter.GetInstance(pdfDocument, fileStream);
                            pdfDocument.Open();
                            pdfDocument.Add(pdfTable);
                            pdfDocument.Close();
                            fileStream.Close();
                            Process.Start(getExactUrl + "ConvertedPdf.pdf");
                        }
                    }
                    else
                    {
                        var newDocument = new Spreadsheet();
                        var sheet = newDocument.Workbook.Worksheets.Add("Converted Excel");
                        int num = 1;
                        foreach (Task<string> item in newList)
                        {
                            var size = item.Result.Split('|');

                            for (int i = 0; i < size.Length; i++)
                            {

                                sheet.Cell(num, i).Value = size[i] ?? " ";

                            }

                            num++;
                        }

                        if (File.Exists(getExactUrl + "ConvertedExcel.xlsx"))
                        {
                            File.Delete(getExactUrl + "ConvertedExcel.xlsx");
                        }
                        newDocument.SaveAs(getExactUrl + "ConvertedExcel.xlsx");
                        newDocument.Close();
                        Process.Start(getExactUrl + "ConvertedExcel.xlsx");
                    }
                }


            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }

            
        }
    }
}
