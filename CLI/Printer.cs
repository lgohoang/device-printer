using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Printing;
using System.Drawing.Printing;
using System.Diagnostics;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace CLI
{
    class Printer
    {
        public delegate void Status(MSD347SDK.Status status, string message, Source source);
        public event Status StatusChanged;

        private BackgroundWorker backgroundWorker;

        bool PrintBySdk;
        bool CheckBySdk;
        bool AutoDetectDriverOnline;

        public enum Source
        {
            System,
            SDK
        }

        MSD347SDK.Status SDKLastStatus = MSD347SDK.Status.PRINTER_IS_OFFLINE;

        public Printer()
        {
            PrintBySdk = false;
            CheckBySdk = false;
            AutoDetectDriverOnline = false;
        }

        public Printer(bool printBySdk, bool checkBySdk, bool autoDetectDriverOnline)
        {
            PrintBySdk = printBySdk;
            CheckBySdk = checkBySdk;
            AutoDetectDriverOnline = autoDetectDriverOnline;
        }

        public void Print()
        {

        }

        public void Start()
        {
            if (backgroundWorker == null)
            {
                backgroundWorker = new BackgroundWorker();
                backgroundWorker.WorkerSupportsCancellation = true;
                backgroundWorker.DoWork += BackgroundWorker_DoWork;
                backgroundWorker.RunWorkerCompleted += BackgroundWorker_RunWorkerCompleted;
            }

            if (!backgroundWorker.IsBusy)
            {
                backgroundWorker.RunWorkerAsync();
            }
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            
        }

        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            PrinterSettings printerSettings = new PrinterSettings();
            LocalPrintServer printerServer = new LocalPrintServer();

            while (!backgroundWorker.CancellationPending)
            {
                //Windows check
                using (var searcher = new ManagementObjectSearcher
                    ("SELECT * FROM WIN32_Printer"))
                {
                    var printers = searcher.Get().Cast<ManagementBaseObject>().Where(p=>p["DriverName"].Equals("MS-D347")).ToList();

                    var defaultPrinterName = printerSettings.PrinterName;

                    var defaultPrinter = printers.Where(p => p["Name"].Equals(defaultPrinterName)).FirstOrDefault();

                    bool Offline = (bool)defaultPrinter["WorkOffline"];

                    if (!Offline)
                    {
                        //Printer online
                        //Print Job != 0 error
                        var queue = printerServer.DefaultPrintQueue;

                        if(queue.NumberOfJobs != 0)
                        {
                            StatusChanged(MSD347SDK.Status.PRINTER_JOBS_QUEUE_NOT_EMPTY, "Jobs: " + queue.NumberOfJobs, Source.System);
                        }

                        //SDK Init
                        if (CheckBySdk)
                        {
                            if (SDKLastStatus.Equals(MSD347SDK.Status.PRINTER_IS_OFFLINE))
                            {
                                try
                                {
                                    //if (MSD347SDK.SetClean() != 0)
                                    //{
                                    //    Debug.WriteLine("SetClean error");
                                    //}

                                    if (MSD347SDK.SetClose() != 0)
                                    {
                                        Debug.WriteLine("SetClose error");
                                    }

                                    if (MSD347SDK.SetUsbportauto() != 0)
                                    {
                                        Debug.WriteLine("SetUsbportauto error");
                                    }

                                    if (MSD347SDK.SetInit() != 0)
                                    {
                                        Debug.WriteLine("SetInit error");
                                    }

                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine("Exception: {0}", ex.ToString());
                                }
                            }

                            //SDK status
                            var statusCode = MSD347SDK.GetStatus();
                            SDKLastStatus = MSD347SDK.GetStatusByCode(statusCode);

                            if (SDKLastStatus.Equals(MSD347SDK.Status.PRINTER_IS_READY))
                            {
                                StatusChanged(MSD347SDK.Status.PRINTER_IS_READY, "", Source.System);
                            }
                            else
                            {
                                StatusChanged(SDKLastStatus, "", Source.SDK);
                            }
                        }
                        

                        Debug.WriteLine("{0} online, status by sdk: {1}", defaultPrinter["Name"], SDKLastStatus);
                    }
                    else
                    {
                        StatusChanged(MSD347SDK.Status.PRINTER_IS_OFFLINE, "", Source.System);

                        Debug.WriteLine("{0} offline", defaultPrinter["Name"]);

                        if (AutoDetectDriverOnline)
                        {
                            foreach (var p in printers)
                            {
                                if (!(bool)p["WorkOffline"])
                                {
                                    var Name = p["Name"].ToString();
                                    SetDefaultPrinter(Name);
                                    StatusChanged(MSD347SDK.Status.PRINTER_DRIVER_CHANGED, "", Source.System);
                                    Debug.WriteLine("{0} is set as the default printer", Name);
                                }
                            }
                        }
                    }  
                }

                Thread.Sleep(500);
            }
        }

        public class MSD347
        {
            private BackgroundWorker backgroundWorker;

            public delegate void MSD347Status(Status status, string message);
            public event MSD347Status StatusChanged;

            public delegate void PrinterStatus(MSD347SDK.Status status, string message);
            public event PrinterStatus PrinterStatusChanged;

            bool AutoUsbport;
            StringBuilder UsbPort;
            int Baudrate;

            public enum Status
            {
                WORKER_STARTED,
                WORKER_STOPED,
                INIT_PRINTER_ERROR,
                SDK_CLEAN_ERROR,
                SDK_CLOSE_ERROR,
                SDK_SET_USB_AUTO_PORT_ERROR,
                SDK_SET_PRINT_PORT_ERROR,
                WINDOWS_PRINTER_IS_OFFLINE,
                WINDOWS_PRINTER_IS_ONLINE,
                WINDOWS_PRINTER_DRIVER_NOT_INSTALL,
                WINDOWS_PRINTER_MANY_DRIVER
            }

            MSD347SDK.Status SDKLastStatus = MSD347SDK.Status.PRINTER_IS_OFFLINE;

            public MSD347()
            {
                AutoUsbport = true;
            }

            public MSD347(string usbPort, int baudrate)
            {
                AutoUsbport = false;
                UsbPort = new StringBuilder(usbPort);
                Baudrate = baudrate;
            }

            public void Start()
            {
                if (backgroundWorker == null)
                {
                    backgroundWorker = new BackgroundWorker();
                    backgroundWorker.WorkerSupportsCancellation = true;
                    backgroundWorker.DoWork += BackgroundWorker_DoWork;
                    backgroundWorker.RunWorkerCompleted += BackgroundWorker_RunWorkerCompleted;
                }

                if (!backgroundWorker.IsBusy)
                {
                    backgroundWorker.RunWorkerAsync();
                }
            }

            public void Stop()
            {
                if(backgroundWorker != null)
                {
                    if (backgroundWorker.IsBusy)
                    {
                        backgroundWorker.CancelAsync();
                    }
                }
            }

            private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
                StatusChanged(Status.WORKER_STOPED, "");
            }

            private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
            {
                StatusChanged(Status.WORKER_STARTED, "");

                while (!backgroundWorker.CancellationPending)
                {
                    //SDK Check
                    if (SDKLastStatus.Equals(MSD347SDK.Status.PRINTER_IS_OFFLINE))
                    {
                        try
                        {
                            if(MSD347SDK.SetClean() != 0)
                            {
                                StatusChanged(Status.SDK_CLEAN_ERROR, "");
                            }

                            if(MSD347SDK.SetClose() != 0)
                            {
                                StatusChanged(Status.SDK_CLOSE_ERROR, "");
                            }

                            if (AutoUsbport)
                            {
                                if (MSD347SDK.SetUsbportauto() != 0)
                                {
                                    StatusChanged(Status.SDK_SET_USB_AUTO_PORT_ERROR, "");
                                }
                            }
                            else
                            {
                                if (MSD347SDK.SetPrintport(UsbPort, Baudrate) != 0)
                                {
                                    StatusChanged(Status.SDK_SET_PRINT_PORT_ERROR, "");
                                }
                            }

                            if(MSD347SDK.SetInit() != 0)
                            {
                                StatusChanged(Status.INIT_PRINTER_ERROR, "");
                            }

                        }
                        catch (Exception ex)
                        {
                            StatusChanged(Status.INIT_PRINTER_ERROR, ex.ToString());
                        }
                    }

                    var statusCode = MSD347SDK.GetStatus();

                    //if (!LastStatus.Equals(MSD347SDK.GetStatusByCode(statusCode)))
                    //{
                    //    PrinterStatusChanged(MSD347SDK.GetStatusByCode(statusCode), "");
                    //    LastStatus = MSD347SDK.GetStatusByCode(statusCode);
                    //}

                    PrinterStatusChanged(MSD347SDK.GetStatusByCode(statusCode), "");
                    SDKLastStatus = MSD347SDK.GetStatusByCode(statusCode);

                    Thread.Sleep(1000);
                }
            }
            
        }

        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool SetDefaultPrinter(string Name);
    }

    
    public static class MSD347SDK
    {
        public enum Status
        {
            PRINTER_IS_READY = 0,
            PRINTER_IS_OFFLINE,
            PRINTER_CALLED_UNMATCHED_LIBRARY,
            PRINTER_HEAD_IS_OPENED,
            PRINTER_CUTTER_IS_NOT_RESET,
            PRINTER_HEAD_TEMP_IS_ABNORMAL,
            PRINTER_DOES_NOT_DETECT_BLACKMARK,
            PRINTER_PAPER_OUT,
            PRINTER_PAPER_LOW,

            PRINTER_DRIVER_CHANGED,
            PRINTER_JOBS_QUEUE_NOT_EMPTY
        }

        public static Status GetStatusByCode(int code)
        {
            return (Status)code;
        }


        //DllImport
        [DllImport("Msprintsdk.dll", EntryPoint = "SetUsbportauto", CharSet = CharSet.Ansi)]
        public static extern unsafe int SetUsbportauto();

        [DllImport("Msprintsdk.dll", EntryPoint = "SetInit", CharSet = CharSet.Ansi)]
        public static extern unsafe int SetInit();

        [DllImport("Msprintsdk.dll", EntryPoint = "SetClean", CharSet = CharSet.Ansi)]
        public static extern unsafe int SetClean();

        [DllImport("Msprintsdk.dll", EntryPoint = "SetClose", CharSet = CharSet.Ansi)]
        public static extern unsafe int SetClose();

        [DllImport("Msprintsdk.dll", EntryPoint = "SetAlignment", CharSet = CharSet.Ansi)]
        public static extern unsafe int SetAlignment(int iAlignment);

        [DllImport("Msprintsdk.dll", EntryPoint = "SetBold", CharSet = CharSet.Ansi)]
        public static extern unsafe int SetBold(int iBold);

        [DllImport("Msprintsdk.dll", EntryPoint = "SetLinespace", CharSet = CharSet.Ansi)]
        public static extern unsafe int SetLinespace(int iLinespace);

        [DllImport("Msprintsdk.dll", EntryPoint = "SetPrintport", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern unsafe int SetPrintport(StringBuilder strPort, int iBaudrate);

        [DllImport("Msprintsdk.dll", EntryPoint = "PrintString", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern unsafe int PrintString(StringBuilder strData);

        [DllImport("Msprintsdk.dll", EntryPoint = "PrintSelfcheck", CharSet = CharSet.Ansi)]
        public static extern unsafe int PrintSelfcheck();

        [DllImport("Msprintsdk.dll", EntryPoint = "GetStatus", CharSet = CharSet.Ansi)]
        public static extern unsafe int GetStatus();

        [DllImport("Msprintsdk.dll", EntryPoint = "PrintFeedline", CharSet = CharSet.Ansi)]
        public static extern unsafe int PrintFeedline(int iLine);

        [DllImport("Msprintsdk.dll", EntryPoint = "PrintCutpaper", CharSet = CharSet.Ansi)]
        public static extern unsafe int PrintCutpaper(int iMode);

        [DllImport("Msprintsdk.dll", EntryPoint = "SetSizetext", CharSet = CharSet.Ansi)]
        public static extern unsafe int SetSizetext(int iHeight, int iWidth);

        [DllImport("Msprintsdk.dll", EntryPoint = "SetSizechinese", CharSet = CharSet.Ansi)]
        public static extern unsafe int SetSizechinese(int iHeight, int iWidth, int iUnderline, int iChinesetype);

        [DllImport("Msprintsdk.dll", EntryPoint = "SetItalic", CharSet = CharSet.Ansi)]
        public static extern unsafe int SetItalic(int iItalic);

        [DllImport("Msprintsdk.dll", EntryPoint = "PrintDiskbmpfile", CharSet = CharSet.Ansi)]
        public static extern unsafe int PrintDiskbmpfile(StringBuilder strData);

        [DllImport("Msprintsdk.dll", EntryPoint = "PrintQrcode", CharSet = CharSet.Ansi)]
        public static extern unsafe int PrintQrcode(StringBuilder strData, int iLmargin, int iMside, int iRound);

        [DllImport("Msprintsdk.dll", EntryPoint = "GetProductinformation", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern unsafe int GetProductinformation(int Fstype, StringBuilder FIDdata);

        [DllImport("Msprintsdk.dll", EntryPoint = "PrintTransmit", CharSet = CharSet.Ansi)]
        public static extern unsafe int PrintTransmit(String strCmd, int iLength);

        [DllImport("Msprintsdk.dll", EntryPoint = "GetTransmit", CharSet = CharSet.Ansi)]
        public static extern unsafe int GetTransmit(string strCmd, int iLength, StringBuilder strRecv, int iRelen);
    }
    
}
