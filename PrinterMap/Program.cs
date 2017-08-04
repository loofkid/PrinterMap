using System;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Printing;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using IWshRuntimeLibrary;

namespace PrinterMap
{
    class Program
    {
        private static string[] printersplit;
        private static string servername;
        private static string sharename;
        private static string printershare;
        private static bool printerExist = false;
        private static Form1 progressBar = new Form1();

        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            printersplit = args[0].TrimStart('z', 'p', 'r', 'i', 'n', 't', 'e', 'r', ':').Split('+');
            servername = printersplit[0];
            sharename = Uri.UnescapeDataString(printersplit[1]);
            printershare = "\\\\" + servername + "\\" + sharename;
            
            BackgroundWorker bg = new BackgroundWorker();
            bg.DoWork += new DoWorkEventHandler(bg_DoWork);
            bg.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bg_RunWorkerCompleted);

            

            foreach (string printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            {
                if (printershare == printer)
                {
                    printerExist = true;
                }
            }
            if (printerExist)
            {
                MessageBox.Show("Printer is already available on your computer.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            } else {
                bg.RunWorkerAsync();
                progressBar.ShowDialog();
            }
            void bg_DoWork(object sender, DoWorkEventArgs e)
            {
                WshNetwork printerAdd = new WshNetwork();
                printerAdd.AddWindowsPrinterConnection(printershare);
            }
            void bg_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
                do
                {
                    foreach (string printer in PrinterSettings.InstalledPrinters)
                    {
                        if (printershare == printer)
                        {
                            printerExist = true;
                        }
                    }
                    Thread.Sleep(500);
                } while (!printerExist);
                progressBar.FormClosed += new FormClosedEventHandler(addComplete);
                progressBar.Close();

                void addComplete(object sender1, FormClosedEventArgs e1)
                {
                    MessageBox.Show("Printer has been successfully added!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
             }
        }
        
    }
}
