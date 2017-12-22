using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Text;
using System.Drawing.Printing;
using System.IO;
using System.Threading;
using Spire.Pdf;
using System.Printing;
using System.Drawing;

namespace CLI
{
    class Program
    {
        static Printer printer;
        static bool printerReady = false;
        static bool userStatusBySDK = false;
        static string printFolder = "ticket";

        static PrinterSettings printerSettings = new PrinterSettings();

        static void Main(string[] args)
        {
            printer = new Printer(true, false, true);
            printer.StatusChanged += Printer_StatusChanged;
            printer.Start();

            if(args.Length > 0)
            {
                printFolder = args[0];
            }

            while (true)
            {
                if (printerReady)
                {
                    var folder = Path.GetFullPath(printFolder);

                    foreach (string file in Directory.GetFiles(folder))
                    {
                        SendToPrinter(Path.GetFullPath(file));
                    }
                }
               
                Thread.Sleep(200);
            }
        }

        private static void Printer_StatusChanged(MSD347SDK.Status status, string message, Printer.Source source)
        {
            //Status by windows device management
            if (source.Equals(Printer.Source.System))
            {
                if (status.Equals(MSD347SDK.Status.PRINTER_IS_READY))
                {
                    printerReady = true;
                }

                if (status.Equals(MSD347SDK.Status.PRINTER_IS_OFFLINE))
                {
                    printerReady = false;
                }

                if (status.Equals(MSD347SDK.Status.PRINTER_DRIVER_CHANGED))
                {
                    //Driver Printer Changed
                }
            }

            //Status by SDK
            if (source.Equals(Printer.Source.SDK) && userStatusBySDK)
            {
                switch (status)
                {
                    case MSD347SDK.Status.PRINTER_PAPER_LOW:
                        //warning paper low
                        break;
                    case MSD347SDK.Status.PRINTER_DOES_NOT_DETECT_BLACKMARK:
                        //warning
                        break;
                    case MSD347SDK.Status.PRINTER_IS_READY:
                        printerReady = true;
                        break;
                    default:
                        printerReady = false;
                        break;
                }
            }

            PrintStatus((int)status, status.ToString(), source.ToString(), "");
        }

        static private void Status()
        {

        }

        private static void PrintStatus(int code, string name, string source, string description)
        {
            var status = new DeviceStatus();
            status.Code = code;
            status.Name = name;
            status.Source = source;
            status.Description = description;

            WriteLine("/status", status.Serialize());
        }

        static private void WriteLine(string uri, string data)
        {
            Console.WriteLine("<<< {0} {1}", uri, data);
        }

        private static void SendToPrinter(string file)
        {
            PdfDocument doc = new PdfDocument();
            try
            {
                doc.LoadFromFile(file);
                doc.PrintFromPage = 0;
                doc.PrintToPage = 1;

                Image i = doc.SaveAsImage(0);
                PrinterSettings settings = new PrinterSettings();
                doc.PrinterName = settings.PrinterName;
                PrintDocument printDoc = doc.PrintDocument;
                var width = printDoc.DefaultPageSettings.PaperSize.Width;
                var height = (int)(((double)printDoc.DefaultPageSettings.PaperSize.Width / (double)i.Width) * (double)i.Height);
                PaperSize ps = new PaperSize("Custom", width, height);
                printDoc.DefaultPageSettings.PaperSize = ps;
                printDoc.EndPrint += (o, e) =>
                {
                    if (printDoc.PrintController.IsPreview)
                        return;
                    File.Delete(file);
                };
                printDoc.Print();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                doc.Close();
            }
        }
    }
}
