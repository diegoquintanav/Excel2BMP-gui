using System;
using System.IO;
using System.Windows;
using Microsoft.Win32;

using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Windows.Forms;

namespace Excel2BMP.Dialogs
{
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {

            Stream myStream = null;

            // opens excel File
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            //openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.RestoreDirectory = false;


            if (openFileDialog.ShowDialog() == true)
            {
                

                var a = new Microsoft.Office.Interop.Excel.Application();

                try
                {
                    if ((myStream = openFileDialog.OpenFile()) != null)
                    {
                        using (myStream)
                        {

                            // creates a new Workbook using the selected Excel File
                            Workbook w = a.Workbooks.Open(openFileDialog.FileName);

                            // gets the root path of the excel File
                            string directoryPath = Path.GetDirectoryName(openFileDialog.FileName);

                            foreach (Worksheet ws in w.Sheets)
                            {
                                string range = ws.PageSetup.PrintArea;
                                ws.Protect(Contents: false);

                                // TODO: Find the first-non blank cell in each sheet to avoid hardcoding
                                // Gets current range
                                Range r = ws.get_Range("B2", System.Type.Missing).CurrentRegion;

                                // Copies the selected range in clipboard as an image
                                r.CopyPicture(XlPictureAppearance.xlScreen, XlCopyPictureFormat.xlBitmap);

                                // saves the Image and displays output
                                Bitmap image = new Bitmap(System.Windows.Forms.Clipboard.GetImage());
                                string BMPName = ws.Name;

                                string imgPath = $@"{directoryPath}\{BMPName}.bmp";

                                txtEditor.Text += $"Saving sheet {BMPName} in {imgPath}...\n";

                                if (File.Exists(imgPath))
                                {
                                    txtEditor.Text += $"File already exists!. Deleting file...\n";
                                    File.Delete(imgPath);
                                }

                                image.Save(imgPath);
                                txtEditor.Text += "Done.\n\n";
                            }
                            //Worksheet ws = w.Sheets["DIA-3"];


                            a.DisplayAlerts = false;

                            // System.Runtime.InteropServices.COMException Excepción de HRESULT: 0x80010105 (RPC_E_SERVERFAULT)
                            //ChartObject chartObj = ws.ChartObjects().Add(r.Left, r.Top, r.Width, r.Height);

                            //chartObj.Activate();
                            //Chart chart = chartObj.Chart;
                            //chart.Paste();
                            //chart.Export(@"C:\Dev\Excel2BMP\Excel2BMP\resources\image.JPG", "JPG");
                            //chartObj.Delete();

                            w.Close(SaveChanges: false);
                        }
                    }

                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("Error: Could not read file: " + ex.Message);
                }
                finally
                {
                    a.Quit();
                }
                

            }



        }
    }
}