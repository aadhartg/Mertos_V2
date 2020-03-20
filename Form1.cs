using Microsoft.Win32;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Auth;
using Microsoft.WindowsAzure.Storage.Blob;
using System;
using System.Collections;
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
using System.Xml.Linq;

namespace WinApp
{
    public partial class Form1 : Form
    {
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.NotifyIcon notifyIcon1;
        private System.Windows.Forms.DataGrid dataGrid1;
        private System.Windows.Forms.MainMenu mainMenu1;
        private System.Windows.Forms.MenuItem menuItem2;
        private System.Windows.Forms.MenuItem menuItem3;


        #region Variables Declaration
        public static int lastProcessTime;
        public static DateTime processStartTime = DateTime.Now;
        public static string appName, prevvalue;
        public static Stack applnames;
        public static Hashtable applhash;
        public static DateTime applfocustime = DateTime.Now;
        public static string appltitle;
        public static Form1 form1;
        public static string tempstr;
        public TimeSpan applfocusinterval;
        public DateTime logintime;
        public DateTime windowLockStartTime;
        public DateTime windowLockEndTime;
        public static bool unLocked { get; private set; }
        public static string Username = Environment.UserName;
        public static string filePath = "C:\\Users\\Public\\Documents\\Mertos";
        public static string fileName = "MertosAppDetails_" + Username + "_" + DateTime.UtcNow.ToShortDateString().Replace('/', '_');
        public static string YesterdayfileName = "MertosAppDetails_" + Username + "_" + DateTime.UtcNow.AddDays(-1).ToShortDateString().Replace('/', '_');
        public string xslFileName = "Mertosappl_xsl.xsl";
        public static bool SystemLock = false;
        public static string currentProcess = "application start$$$!!!Mertos%%%###" + DateTime.Now.ToString();

        private CloudBlobClient _client;
        private CloudBlobContainer _container;
        private Timer timer2;
        private string CONTAINER_NAME = "watcher";

        #endregion

        public string Get(string name)
        {
            // Retrieve reference to a blob by filename, e.g. "photo1.jpg".
            var blob = _container.GetBlockBlobReference(name);
            var blobInfoString = $"Block blob with name '{blob.Name}', " +
                        $"content type '{blob.Properties.ContentType}', " +
                        $"size '{blob.Properties.Length}', " +
                        $"and URI '{blob.Uri}'";
            return blobInfoString;
        }

        public static void CheckFileExists()
        {
            var user = Username;
            bool directoryExist = Directory.Exists(filePath);
            if (!directoryExist)
            {
                Directory.CreateDirectory(filePath);
                if (File.Exists(Path.Combine(filePath, YesterdayfileName + ".xml")))
                {
                    File.Move(Path.Combine(filePath, YesterdayfileName + ".xml"), Path.Combine(filePath, YesterdayfileName + ".xml.bak"));
                }
            }
            var oldpath = Path.Combine(filePath, fileName);
            if (File.Exists(oldpath + ".xml"))
            {
                if (Directory.GetFiles(filePath, fileName + "*.bak").Length != 0)
                {
                    var backupCount = Directory.GetFiles(filePath, fileName + "*xml.bak").Length;

                    var newPath = (oldpath + "_" + Convert.ToString(backupCount + 1) + ".xml.bak");
                    File.Move(oldpath + ".xml", newPath);
                    File.Create(oldpath + ".xml").Close();
                    StreamWriter writer = null;
                    FileStream stream = new FileStream(Path.Combine(filePath, fileName + ".xml"), FileMode.Open);
                    using (writer = new System.IO.StreamWriter(stream, Encoding.UTF8))
                    {
                        writer.Write("<?xml version=\"1.0\"?>");
                        writer.WriteLine("");
                        writer.WriteLine("<?xml-stylesheet type=\"text/xsl\" href=\"appl_xsl.xsl\"?>");
                        writer.WriteLine("");
                        writer.WriteLine("<ApplDetails>");
                        writer.WriteLine("</ApplDetails>");
                    }
                }
                else
                {
                    oldpath = oldpath + ".xml";
                    var newPath = (filePath + "\\" + fileName + "_" + "1" + ".xml.bak");
                    File.Move(oldpath, newPath);
                    File.Create(oldpath).Close();
                    StreamWriter writer = null;
                    FileStream stream = new FileStream(Path.Combine(filePath, fileName + ".xml"), FileMode.Open);
                    using (writer = new System.IO.StreamWriter(stream, Encoding.UTF8))
                    {
                        writer.Write("<?xml version=\"1.0\"?>");
                        writer.WriteLine("");
                        writer.WriteLine("<?xml-stylesheet type=\"text/xsl\" href=\"appl_xsl.xsl\"?>");
                        writer.WriteLine("");
                        writer.WriteLine("<ApplDetails>");
                        writer.WriteLine("</ApplDetails>");
                    }
                }
            }
            else
            {
                File.Create(oldpath + ".xml").Close();
                StreamWriter writer = null;
                FileStream stream = new FileStream(Path.Combine(filePath, fileName + ".xml"), FileMode.Open);
                using (writer = new System.IO.StreamWriter(stream, Encoding.UTF8))
                {
                    writer.Write("<?xml version=\"1.0\"?>");
                    writer.WriteLine("");
                    writer.WriteLine("<?xml-stylesheet type=\"text/xsl\" href=\"appl_xsl.xsl\"?>");
                    writer.WriteLine("");
                    writer.WriteLine("<ApplDetails>");
                    writer.WriteLine("</ApplDetails>");
                }
            }
        }
        public Form1()
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent1();

            //
            // TODO: Add any constructor code after InitializeComponent call
            //
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent1()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.timer2 = new System.Windows.Forms.Timer(this.components);
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.dataGrid1 = new System.Windows.Forms.DataGrid();
            this.mainMenu1 = new System.Windows.Forms.MainMenu(this.components);
            this.menuItem2 = new System.Windows.Forms.MenuItem();
            this.menuItem3 = new System.Windows.Forms.MenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid1)).BeginInit();
            this.SuspendLayout();
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Interval = 1000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // timer2
            // 
            this.timer2.Enabled = true;
            this.timer2.Interval = 60000;
            this.timer2.Tick += new System.EventHandler(this.timer2_Tick);
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon1.Icon")));
            this.notifyIcon1.Text = "notifyIcon1";
            this.notifyIcon1.Visible = true;
            // 
            // dataGrid1
            // 
            this.dataGrid1.DataMember = "";
            this.dataGrid1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGrid1.GridLineColor = System.Drawing.SystemColors.Highlight;
            this.dataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dataGrid1.Location = new System.Drawing.Point(0, 0);
            this.dataGrid1.Name = "dataGrid1";
            this.dataGrid1.PreferredColumnWidth = 164;
            this.dataGrid1.Size = new System.Drawing.Size(536, 461);
            this.dataGrid1.TabIndex = 0;
            // 
            // menuItem2
            // 
            this.menuItem2.Index = -1;
            this.menuItem2.Shortcut = System.Windows.Forms.Shortcut.F5;
            this.menuItem2.Text = "&Refresh";
            // 
            // menuItem3
            // 
            this.menuItem3.Index = -1;
            this.menuItem3.Text = "E&xit";
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(536, 461);
            this.Controls.Add(this.dataGrid1);
            this.MaximizeBox = false;
            this.Menu = this.mainMenu1;
            this.Name = "Form1";
            this.ShowInTaskbar = false;
            this.Text = "Win Appl Watcher";
            this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid1)).EndInit();
            this.ResumeLayout(false);

        }
        #endregion
        [STAThread]
        static void Main()
        {
            CheckFileExists();
            SystemEvents.SessionSwitch += new SessionSwitchEventHandler(SystemEvents_SessionSwitch);
            applnames = new Stack();
            applhash = new Hashtable();
            form1 = new Form1();
            Application.Run(form1);
        }

        static void SystemEvents_SessionSwitch(object sender, SessionSwitchEventArgs e)
        {
            if (e.Reason == SessionSwitchReason.SessionLock)
            {
                //I left my desk
                SystemLock = true;
            }
            else if (e.Reason == SessionSwitchReason.SessionUnlock)
            {
                //I returned to my desk
                SystemLock = false;
            }
        }

        private void timer2_Tick(object sender, System.EventArgs e)
        {
            try
            {
                UploadFileInBlob();
            }
            catch (Exception ex)
            {
            }
        }

        private void UploadFileInBlob()
        {
            try
            {
                StorageCredentials credentials = new StorageCredentials("mertos", "4nAyKHDRse3ggXRpdNVMHVuXDJcqJlrG/p97QNuT4hC4R3859GuA3www6wMIV8d4jDsPrCSWfaOfpcqvX1JoCw==");
                CloudStorageAccount storageAccount = new CloudStorageAccount(credentials, true);
                _client = storageAccount.CreateCloudBlobClient();
                _container = _client.GetContainerReference(CONTAINER_NAME);
                var result = _container.CreateIfNotExistsAsync();
                result.Wait();
                var aa = result.Result;
                if (aa)
                {
                    _container.SetPermissions(new
                       BlobContainerPermissions
                    {
                        PublicAccess =
                       BlobContainerPublicAccessType.Blob
                    });
                }

                string xmlfile = Path.Combine(filePath, fileName + ".xml");

                CloudBlockBlob cloudBlockBlob = _container.GetBlockBlobReference("mertos Data/" + Username + "/" + fileName + ".xml");
                //cloudBlockBlob.Properties.ContentType = file.;
                cloudBlockBlob.UploadFromFileAsync(xmlfile);
                //using (Stream file = System.IO.File.OpenRead(xmlfile))
                //{
                //    cloudBlockBlob.UploadFromStreamAsync(file);
                //}
            }
            catch (Exception ex)
            {
            }
        }

        private void timer1_Tick(object sender, System.EventArgs e)
        {
            //This is used to monitor and save active application's  details in Hashtable for future saving in xml file...
            try
            {

                bool dirExist = Directory.Exists(filePath);
                if (!dirExist)
                {
                    Directory.CreateDirectory(filePath);
                }
                var oldpath = Path.Combine(filePath, fileName);
                if (!File.Exists(oldpath + ".xml"))
                {
                    File.Create(oldpath + ".xml").Close();
                    StreamWriter writer = null;
                    FileStream stream = new FileStream(Path.Combine(filePath, fileName + ".xml"), FileMode.Open);
                    using (writer = new System.IO.StreamWriter(stream, Encoding.UTF8))
                    {
                        writer.Write("<?xml version=\"1.0\"?>");
                        writer.WriteLine("");
                        writer.WriteLine("<?xml-stylesheet type=\"text/xsl\" href=\"appl_xsl.xsl\"?>");
                        writer.WriteLine("");
                        writer.WriteLine("<ApplDetails>");
                        writer.WriteLine("</ApplDetails>");
                    }
                }
                if (prevvalue == null)
                {
                    prevvalue = "application start$$$!!!Mertos%%%###" + DateTime.Now.ToString();
                }
                // bool isNewAppl = false;
                IntPtr hwnd = APIFuncs.getforegroundWindow();
                Int32 pid = APIFuncs.GetWindowProcessID(hwnd);
                Process p = Process.GetProcessById(pid);

                var IdleTime = APIFuncs.IdleTimeFinder.GetIdleTime();

                if (IdleTime > 30000 && SystemLock == false)
                {
                    appName = "Idle";
                    appltitle = "Idle";
                }
                else if (SystemLock == true)
                {
                    appName = "System Lock";
                    appltitle = "System Lock";
                }
                else if (appltitle == null && appName == null)
                {
                    appltitle = "application start";
                    appName = "Mertos";
                }
                else if (appltitle == null || appltitle == "")
                {
                    appltitle = appName;
                    currentProcess = appltitle + "$$$!!!" + appName + "%%%###" + Convert.ToString(processStartTime);
                }
                else
                {
                    appName = p.ProcessName;
                    appltitle = APIFuncs.ActiveApplTitle().Trim().Replace("\0", "");
                }
                if (prevvalue != (appltitle + "$$$!!!" + appName))
                {
                    prevvalue = appltitle + "$$$!!!" + appName;
                    SaveandShowDetails();
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("On Timer ----" + ex.Message);
            }
        }

        public static void WriteToFile(string Message)
        {
            try
            {
                string path = "e:\\Logs";
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                string filepath = "e:\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
                if (!File.Exists(filepath))
                {
                    // Create a file to write to.   
                    using (StreamWriter sw = File.CreateText(filepath))
                    {
                        sw.WriteLine(Message);
                    }
                }
                else
                {
                    using (StreamWriter sw = File.AppendText(filepath))
                    {
                        sw.WriteLine(Message);
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
        public static string GetTotalTime(DateTime StartTime, DateTime processEndTime)
        {
            //  TimeSpan time = StartTime.Subtract(processEndTime).Negate();
            TimeSpan diffTime = (processEndTime - StartTime);
            // int seconds =  Convert.ToInt32(time.TotalSeconds);
            //if(seconds > 60)
            //{
            //    return (seconds / 60) + " Minutes";
            //}
            //else
            //{
            //    return seconds + " Second";
            //}
            return string.Concat(diffTime.Hours + " Hours " + diffTime.Minutes + " Mins " + Convert.ToString(diffTime.Seconds + 1) + " secs");
        }
        private void SaveandShowDetails()
        {
            //This is used to save contents of hashtable in a xml file....
            // StreamWriter writer=null;
            try
            {
                string processEndTime = DateTime.Now.ToString();
                string totalTime = GetTotalTime(processStartTime, Convert.ToDateTime(processEndTime));
                string[] separator = { "$$$!!!", "%%%###" };
                var data = currentProcess.Split(separator, 5, StringSplitOptions.RemoveEmptyEntries);
                XDocument xd = XDocument.Load(Path.Combine(filePath, fileName + ".xml"));

                xd.Element("ApplDetails").Add(
                    new XElement("Application_Info",
                    new XElement("ProcessName", data[0]),
                    new XElement("ApplicationName", data[1]),
                    new XElement("TotalTime", totalTime),
                    new XElement("StartTime", data[2]),
                    new XElement("endTime", processEndTime)));
                xd.Save((Path.Combine(filePath, fileName + ".xml")));
                processStartTime = DateTime.Now;
                currentProcess = appltitle + "$$$!!!" + appName + "%%%###" + processStartTime.ToString();
            }
            catch (Exception ex)
            {
                // MessageBox.Show("On SaveandShowDetails ---" + ex.Message);
            }
        }
        private void Form1_Load(object sender, System.EventArgs e)
        {
            form1.Visible = false;
            notifyIcon1.Text = "Mertos is in Invisible Mode";
            logintime = DateTime.Now;
            form1.Text = "Login Time is at :" + DateTime.Now.ToLongTimeString();
        }
    }
}
