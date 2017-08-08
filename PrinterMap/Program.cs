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
using System.Net.NetworkInformation;
using System.Printing;

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
        private static PingReply pingReply;

        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            printersplit = args[0].TrimStart('z', 'p', 'r', 'i', 'n', 't', 'e', 'r', ':').Split('+');
            servername = printersplit[0];
            sharename = Uri.UnescapeDataString(printersplit[1]);
            printershare = "\\\\" + servername + "\\" + sharename;
                      
            
            Ping ping = new Ping();
            try { 
                pingReply = ping.Send(servername);
            } catch
            {
                MessageBox.Show("Print server is not available.\nPlease make sure you can access Zeon's network.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (pingReply.Status == IPStatus.Success)
            {
                PrintServer printServer = new PrintServer(@"\\" + servername);
                try
                {
                    PrintQueue printQueue = printServer.GetPrintQueue(printershare);
                } catch
                {
                    MessageBox.Show("Printer cannot be found on print server.\nPlease verify that printer exists.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                BackgroundWorker bg = new BackgroundWorker();
                bg.DoWork += new DoWorkEventHandler(bg_DoWork);
                bg.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bg_RunWorkerCompleted);

                foreach (string printer in PrinterSettings.InstalledPrinters)
                {
                    if (printershare == printer)
                    {
                        printerExist = true;
                    }
                }
                if (printerExist)
                {
                    MessageBox.Show("Printer is already available on your computer.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
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
            } else
            {
                MessageBox.Show("Print server is not available.\nPlease make sure you can access Zeon's network.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
    }
}
