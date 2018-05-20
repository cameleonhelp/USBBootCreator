using Microsoft.VisualBasic.FileIO;
using Microsoft.Win32;
using System;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace USBBootCreator
{
    public partial class Form1 : Form
    {
        private Thread t;
        public int lastVal = 0;
        public string HKLMUsbCreator = "HKEY_CURRENT_USER\\SOFTWARE\\e.SNCF\\USBCreator\\Settings";
        public static string logfile = @"C:\e.SNCF\logs\usbcreator.log";
        public enum logtype { Information, Success, Warning, Error, Failure };
        public Task[] arrayTasks = { };

        public void Log(int iLogType,string logMessage)
        {
            using (StreamWriter w = File.AppendText(logfile))
            {
                w.WriteLine(DateTime.Now.ToString("") + ";" + Enum.GetName(typeof(logtype),iLogType) + ";" + logMessage);
            }
        }

        public Form1()
        {
            t = new Thread(new ThreadStart(StartSplashScreen));
            InitializeComponent();
        }

        public void StartSplashScreen()
        {
            Application.Run(new SplashScreen());
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        public bool IsWindows10()
        {
            var reg = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion");

            string productName = (string)reg.GetValue("ProductName");

            return productName.StartsWith("Windows 10");
        }

        public bool IsWindows7()
        {
            var reg = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion");

            string productName = (string)reg.GetValue("ProductName");

            return productName.StartsWith("Windows 7");
        }

        public void readRegParam()
        {
            try
            {
                textBox1.Text = (string)Registry.GetValue(HKLMUsbCreator, "MasterPath", Application.StartupPath + @"\Masters");
                //if (textBox1.Text == "")
                //    textBox1.Text = Application.StartupPath + @"\Masters";
                //comboBox3.Text = (string)Registry.GetValue(HKLMUsbCreator, "lastMasterUsed", "");
                string trackPosition = (string)Registry.GetValue(HKLMUsbCreator, "PartSize", "4");
                trackBar1.Value = Convert.ToInt32(trackPosition);
                
                Registry.SetValue(HKLMUsbCreator, "Version", Application.ProductVersion, RegistryValueKind.String);
                Registry.SetValue(HKLMUsbCreator, "MasterPath", textBox1.Text, RegistryValueKind.String);
                Registry.SetValue(HKLMUsbCreator, "lastMasterUsed", comboBox3.Text, RegistryValueKind.String);
                Registry.SetValue(HKLMUsbCreator, "PartSize", trackBar1.Value, RegistryValueKind.String);
            }
            catch (Exception ex)  
            {
                Log(3, ex.Message);
                ex.GetType();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            t.Start();
            if (File.Exists(logfile))
            {
                File.Delete(logfile);
            }
            if (!Directory.Exists(@"C:\e.SNCF\logs"))
            {
                Directory.CreateDirectory(@"C:\e.SNCF\logs");
            }
            Log(0, "---DEBUT CHARGEMENT DU PROGRAMME---");
            var versionInfo = FileVersionInfo.GetVersionInfo(Assembly.GetEntryAssembly().Location);
            var copyright = versionInfo.LegalCopyright;
            label8.Text = "Version " + Application.ProductVersion + " - " + copyright;
            Log(0,label8.Text);
            this.Cursor = Cursors.WaitCursor;
            readRegParam();
            Application.DoEvents();
            label3.Text = "Détection du système d'exploitation ...";
            Log(0,label3.Text);
            Application.DoEvents();
            if (IsWindows10()) {
                Log(0, "Microsoft Windows 10 " + GetWindowsVersion());
                Application.DoEvents();
                label3.Text = "Détection de l'espace libre sur votre disque C: ...";
                Log(0,label3.Text);
                Application.DoEvents();
                if (GetTotalFreeSpace("C:\\") > 20)
                {
                    Application.DoEvents();
                    label3.Text = "Remplissage des disques USB avec une partition ...";
                    Log(0, label3.Text);
                    Application.DoEvents();
                    SetAllUSBDrive();
                    Application.DoEvents();
                    label3.Text = "Remplissage des ISO de masters ...";
                    Log(0, label3.Text);
                    Application.DoEvents();
                    ListAllMasters();
                    Application.DoEvents();
                    label3.Text = "Remplissage des partitions de Boot ...";
                    Log(0, label3.Text);
                    Application.DoEvents();
                    SetAllUSBDriveFat32();
               } else
                {
                    Application.DoEvents();
                    label3.Text = "En attente ...";
                    Application.DoEvents();
                    this.Cursor = Cursors.Default;
                    MessageBox.Show("Votre espace libre sur C: est inférieure à 20 Go\n\rce qui est l'espace minimum requis pour préparer un disque USB bootable.", "Espace insuffisant", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    Log(2, "Espace libre sur C: inférieure à 20Go");
                    this.Close();
                }
            } else if (IsWindows7()) {
                // test des pré-requis : 
                Log(0, "Microsoft Windows 7 " + GetWindowsVersion());
                Application.DoEvents();
                label3.Text = "Détection de la présence du fichier WinCDEmu.exe : ...";
                Log(0, label3.Text);
                Application.DoEvents();
                if (File.Exists(Application.StartupPath + "\\WinCDEmu.exe"))
                {
                    Application.DoEvents();
                    label3.Text = "Détection de l'espace libre sur votre disque C: ...";
                    Log(0, label3.Text);
                    Application.DoEvents();
                    if (GetTotalFreeSpace("C:\\") > 20)
                    {
                        Application.DoEvents();
                        label3.Text = "Remplissage des disques USB avec une partition ...";
                        Log(0, label3.Text);
                        Application.DoEvents();
                        SetAllUSBDrive();
                        Application.DoEvents();
                        label3.Text = "Remplissage des ISO de masters ...";
                        Log(0, label3.Text);
                        Application.DoEvents();
                        ListAllMasters();
                        Application.DoEvents();
                        label3.Text = "Remplissage des partitions de Boot ...";
                        Log(0, label3.Text);
                        Application.DoEvents();
                        SetAllUSBDriveFat32();
                    }
                    else
                    {
                        Application.DoEvents();
                        label3.Text = "En attente ...";
                        Application.DoEvents();
                        this.Cursor = Cursors.Default;
                        MessageBox.Show("Votre espace libre sur C: est inférieure à 20 Go\n\rce qui est l'espace minimum requis pour préparer un disque USB bootable.", "Espace insuffisant", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        Log(2, "Espace libre sur C: inférieure à 20Go");
                        this.Close();
                    }
                } else
                {
                    Application.DoEvents();
                    label3.Text = "En attente ...";
                    Application.DoEvents();
                    this.Cursor = Cursors.Default;
                    MessageBox.Show("Le fichier WinCDEmu.exe est manquant.\r\nVeuillez corriger ce problème.", "Fichier manquant", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    Log(4, "WinCDEmu manquant");
                    this.Close();
                }
            }
            else {
                Application.DoEvents();
                label3.Text = "En attente ...";
                Application.DoEvents();
                this.Cursor = Cursors.Default;
                MessageBox.Show("Windows 7 ou Windows 10 est un pré-requis à l'utilisation de ce programme","Système non compatible", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                Log(4, "OS non suporte");
                this.Close();
            }
            Application.DoEvents();
            label3.Text = "En attente ...";
            Application.DoEvents();
            this.Cursor = Cursors.Default;
            Log(0, "---FIN CHARGEMENT DU PROGRAMME---");
            t.Abort();
        }

        public long GetTotalFreeSpace(string driveName)
        {
            foreach (DriveInfo drive in DriveInfo.GetDrives())
            {
                if (drive.IsReady && drive.Name == driveName)
                {
                    return (drive.TotalFreeSpace/ (1024*1024*1024));
                }
            }
            return -1;
        }

        public string OtherLetterPartition(string logicalDiskId)
        {
            var deviceId = string.Empty;
            string otherLetter = null;

            var query = "ASSOCIATORS OF {Win32_LogicalDisk.DeviceID='" + logicalDiskId + "'} WHERE AssocClass = Win32_LogicalDiskToPartition";
            var queryResults = new ManagementObjectSearcher(query);
            var partitions = queryResults.Get();

            foreach (var partition in partitions)
            {
                query = "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" + partition["DeviceID"] + "'} WHERE AssocClass = Win32_DiskDriveToDiskPartition";
                queryResults = new ManagementObjectSearcher(query);
                var drives = queryResults.Get();


                foreach (var drive in drives)
                {
                    var letter = drive["Caption"];
                    deviceId = drive["DeviceID"].ToString();
                }
            }

            foreach (ManagementObject partition in new ManagementObjectSearcher(
                "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='" + deviceId
                + "'} WHERE AssocClass = Win32_DiskDriveToDiskPartition").Get())
            {
                foreach (ManagementObject disk in new ManagementObjectSearcher(
                            "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='"
                                + partition["DeviceID"]
                                + "'} WHERE AssocClass = Win32_LogicalDiskToPartition").Get())
                {
                    var letter = disk["Name"];
                    if (letter.ToString() != logicalDiskId)
                    {
                        otherLetter = letter.ToString();
                    }
                }
            }

            return otherLetter;
        }


        public void SetAllUSBDriveFat32()
        {
            comboBox1.Enabled = false;
            comboBox1.Items.Clear();
            comboBox1.Text = "";
            button2.Enabled = false;
            button2.BackColor = Color.FromArgb(224, 222, 216);

            this.comboBox1.Items.Add("");
            foreach (ManagementObject drive in new ManagementObjectSearcher("select * from Win32_DiskDrive where InterfaceType='USB'").Get())
            {
                if (Convert.ToString(drive["MediaType"]).ToLower().Contains("external") || Convert.ToString(drive["MediaType"]).ToLower().Contains("removable"))
                {
                    if (Convert.ToString(drive["Partitions"]) == "2")
                    {
                        Application.DoEvents();
                        comboBox1.Enabled = true;
                        button2.Enabled = true;
                        button2.BackColor = Color.FromArgb(0, 136, 206);
                        foreach (ManagementObject partition in new ManagementObjectSearcher("ASSOCIATORS OF {Win32_DiskDrive.DeviceID='" + drive["DeviceID"] + "'} WHERE AssocClass = Win32_DiskDriveToDiskPartition").Get())
                        {
                            // associate partitions with logical disks (drive letter volumes)
                            Application.DoEvents();
                            foreach (ManagementObject mdisk in new ManagementObjectSearcher("ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" + partition["DeviceID"] + "'} WHERE AssocClass =Win32_LogicalDiskToPartition").Get())
                            {
                                if (mdisk["FileSystem"].ToString() == "FAT32")
                                {
                                    Application.DoEvents();
                                    this.comboBox1.Items.Add(mdisk["Name"]);
                                    Log(0, "Ajout de la parttion de BOOT : " + mdisk["Name"]);
                                }
                            }
                        }
                    }
                }
            }
            comboBox1.SelectedIndex = 0;
        }

        public void SetAllUSBDrive()
        {
            comboBox4.Items.Clear();
            comboBox4.Text = "";
            comboBox4.Enabled = false;
            button1.Enabled = false;
            button1.BackColor = Color.FromArgb(224, 222, 216);
            trackBar1.Enabled = false;

            this.comboBox4.Items.Add("");
            foreach (ManagementObject drive in new ManagementObjectSearcher("select * from Win32_DiskDrive where InterfaceType='USB'").Get())
            {
                if (Convert.ToString(drive["MediaType"]).ToLower().Contains("external") || Convert.ToString(drive["MediaType"]).ToLower().Contains("removable"))
                {
                    if (Convert.ToString(drive["Partitions"]) == "1")
                    {
                        Application.DoEvents();
                        foreach (ManagementObject partition in new ManagementObjectSearcher("ASSOCIATORS OF {Win32_DiskDrive.DeviceID='" + drive["DeviceID"] + "'} WHERE AssocClass = Win32_DiskDriveToDiskPartition").Get())
                        {
                            // associate partitions with logical disks (drive letter volumes)
                            foreach (ManagementObject mdisk in new ManagementObjectSearcher("ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" + partition["DeviceID"] + "'} WHERE AssocClass =Win32_LogicalDiskToPartition").Get())
                            {
                                Application.DoEvents();
                                this.comboBox4.Items.Add(mdisk["Name"]);
                                comboBox4.Enabled = true;
                                Log(0, "Ajout du lecteur : " + mdisk["Name"]);
                            }
                        }
                    }
                }
            }
            comboBox4.SelectedIndex = 0;
        }

        public async void button1_Click(object sender, EventArgs e)
        {
            Log(0, "---DEBUT DU PARTIONNEMENT---");
            Log(0, "Partitionnement du disque : " + comboBox4.Text);
            this.Cursor = Cursors.WaitCursor;
            //Partionnement du disque sélectionné 16Go en FAT32 et reste en NTFS
            Application.DoEvents();
            label3.Text = "Identification du disque ...";
            Log(0, label3.Text);
            Application.DoEvents();
            int idx = GetDiskIndex(comboBox4.Text);
            Application.DoEvents();
            label3.Text = "Disque identifié avec l'index " + idx + " ...";
            Log(0, label3.Text);
            Application.DoEvents();
            if (idx != -1)
            {
                comboBox4.Enabled = false;
                button1.Enabled = false;
                button1.BackColor = Color.FromArgb(224, 222, 216);
                Application.DoEvents();
                label3.Text = "Préparation du disque en cours ...";
                Log(0, label3.Text);
                Application.DoEvents();
                await Task.Run(() =>
                {
                    DiskPart(idx.ToString(), comboBox4.Text + "\\");
                });
                comboBox4.Enabled = true;
                button1.Enabled = true;
                button1.BackColor = Color.FromArgb(0, 136, 206);
                Application.DoEvents();
                label3.Text = "Disque prêt ...";
                Log(0, label3.Text);
                Application.DoEvents();
            }
            comboBox4.Text = "";
            comboBox4.Items.Clear();
            Application.DoEvents();
            label3.Text = "Mise à jour de la liste des partition de Boot ...";
            Log(0, label3.Text);
            Application.DoEvents();
            SetAllUSBDrive();
            SetAllUSBDriveFat32();
            label3.Text = "Mise à jour de la liste des masters ...";
            Log(0, label3.Text);
            ListAllMasters();
            this.Cursor = Cursors.Default;
            Application.DoEvents();
            label3.Text = "Vous pouvez copier vos ISO dans le dossier \"Masters\".";
            Log(0, label3.Text);
            Application.DoEvents();
            Log(0, "---FIN DU PARTIONNEMENT---");
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox4.Text == "")
            {
                button1.Enabled = false;
                button1.BackColor = Color.FromArgb(224, 222, 216);
                trackBar1.Enabled = false;
            } else {
                if (GetTotalSize(comboBox4.Text + "\\") > 32)
                {
                    button1.Enabled = true;
                    button1.BackColor = Color.FromArgb(0, 136, 206);
                    trackBar1.Enabled = true;
                }
                else
                {
                    button1.Enabled = false;
                    button1.BackColor = Color.FromArgb(224, 222, 216);
                    trackBar1.Enabled = false;
                    MessageBox.Show("La capacité du lecteur " + comboBox4.Text + " n'est pas suffisante.\r\nVotre périphérique doit avoir une capacité minimum de 64Go.", "Capacité insuffisante", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            comboBox4.Text = "";
            comboBox4.Items.Clear();
            SetAllUSBDrive();
        }

        private int GetDiskIndex(string driveLetter)
        {
            driveLetter = driveLetter.TrimEnd('\\');

            ManagementScope scope = new ManagementScope(@"\root\cimv2");
            var drives = new ManagementObjectSearcher(scope, new ObjectQuery("select * from Win32_DiskDrive")).Get();
            foreach (var drive in drives)
            {

                var partitions = new ManagementObjectSearcher(scope, new ObjectQuery("ASSOCIATORS OF {Win32_DiskDrive.DeviceID='" + drive["DeviceID"] + "'} WHERE AssocClass = Win32_DiskDriveToDiskPartition")).Get();
                foreach (var partition in partitions)
                {
                    var logicalDisks = new ManagementObjectSearcher(scope, new ObjectQuery("ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" + partition["DeviceID"] + "'} WHERE AssocClass = Win32_LogicalDiskToPartition")).Get();
                    foreach (var logicalDisk in logicalDisks)
                    {
                        if (logicalDisk["DeviceId"].ToString() == driveLetter) return Convert.ToInt32(partition["DiskIndex"]);
                    }
                }

            }

            return -1;
        }

        public async Task<string> MountISO(string fIso)
        {
            string result = null;
            try
            {
                string scriptFilePath = @"C:\psScript.ps1";

                if (File.Exists(scriptFilePath))
                {
                    // Note: if the file doesn't exist, no exception is thrown    
                    File.Delete(scriptFilePath); // Delete the script, if it exists    
                }

                // Create the script to resize F and make U    
                File.AppendAllText(scriptFilePath,
                string.Format(
                    "$mountResult = Mount-DiskImage -ImagePath \"" + fIso + "\" -PassThru\r\n" +
                    "$driveLetter = ($mountResult | Get-Volume).DriveLetter\r\n" +
                    "Write-Host $driveLetter"
                    )
                ); // And exit    

                int exitcode = 0;
                result = await Task.Run(() =>
                {
                    return ExecutePSCommand("Powershell.exe" + " -executionpolicy bypass -File \"" + scriptFilePath + "\"", ref exitcode);
                });
                File.Delete(scriptFilePath); // Delete the script file    
                if (exitcode > 0)
                {
                    result = null;
                }
                return result;
            }
            catch (Exception ex)
            {
                Log(3, ex.Message);
                return null;
                ex.GetType();
            }
        }

        public async Task<string> MountISO7(string fIso)
        {
            string result = null;
            var unusedDrives = "CDEFGHIJKLMNOPQRSTUVWXYZ".Except(from d in DriveInfo.GetDrives() select d.Name.First());
            var firstUnusedDrive = unusedDrives.First();
            try
            {
                string scriptFilePath = @"C:\psScript.ps1";

                if (File.Exists(scriptFilePath))
                {
                    // Note: if the file doesn't exist, no exception is thrown    
                    File.Delete(scriptFilePath); // Delete the script, if it exists    
                }

                // Create the script to resize F and make U    
                File.AppendAllText(scriptFilePath,
                string.Format(
                    Application.StartupPath + "WinCDEmu.exe /install\r\n" +
                    "Start-Sleep -s 5\r\n" +
                    Application.StartupPath + "WinCDEmu.exe \"" + fIso + "\" " + firstUnusedDrive.ToString() + " /wait"
                    )
                ); // And exit    

                int exitcode = 0;
                result = await Task.Run(() =>
                {
                    return ExecutePSCommand("Powershell.exe" + " -executionpolicy bypass -File \"" + scriptFilePath + "\"", ref exitcode);
                });
                //File.Delete(scriptFilePath); // Delete the script file    
                if (exitcode > 0)
                {
                    result = null;
                }
                DriveInfo d = new DriveInfo(firstUnusedDrive.ToString());
                int i = 0;
                do
                {
                    i++;
                } while (!d.IsReady);
                return firstUnusedDrive.ToString();
            }
            catch (Exception ex)
            {
                Log(3, ex.Message);
                return null;
                ex.GetType();
            }
        }

        public async void UnMountISO(string fIso)
        {
            string result = null;
            try
            {
                string scriptFilePath = @"C:\psScript.ps1";

                if (File.Exists(scriptFilePath))
                {
                    // Note: if the file doesn't exist, no exception is thrown    
                    File.Delete(scriptFilePath); // Delete the script, if it exists    
                }

                // Create the script to resize F and make U    
                File.AppendAllText(scriptFilePath,
                string.Format(
                    "$mountResult = Dismount-DiskImage -ImagePath \"" + fIso + "\" \r\n"
                    )
                ); // And exit    

                int exitcode = 0;
                result = await Task.Run(() =>
                {
                    return ExecuteCmdCommand("Powershell.exe" + " -executionpolicy bypass -File \"" + scriptFilePath + "\"", ref exitcode);
                });
                File.Delete(scriptFilePath); // Delete the script file    
                if (exitcode > 0)
                {
                    Log(2, "UNmountISO result : " + "[" + exitcode + "]" + result);
                    result = null;
                }
                Log(0, "UNmountISO result : " + "[" + exitcode + "]" + result);
            }
            catch (Exception ex)
            {
                Log(3, ex.Message);
                ex.GetType();
            }
        }

        public async void UnMountISO7(string destDrive)
        {
            string result = null;
            try
            {
                string scriptFilePath = @"C:\psScript.ps1";

                if (File.Exists(scriptFilePath))
                {
                    // Note: if the file doesn't exist, no exception is thrown    
                    File.Delete(scriptFilePath); // Delete the script, if it exists    
                }

                // Create the script to resize F and make U    
                File.AppendAllText(scriptFilePath,
                string.Format(
                    Application.StartupPath + "WinCDEmu.exe /unmount " + destDrive + "\r\n" +
                    "Start-Sleep -s 5\r\n" +
                    Application.StartupPath + "WinCDEmu.exe /uninstall"
                    )
                ); // And exit    

                int exitcode = 0;
                result = await Task.Run(() =>
                {
                    return ExecuteCmdCommand("Powershell.exe" + " -executionpolicy bypass -File \"" + scriptFilePath + "\"", ref exitcode);
                });
                File.Delete(scriptFilePath); // Delete the script file    
                if (exitcode > 0)
                {
                    Log(2, "UNmountISO result : " + "[" + exitcode + "]" + result);
                    result = null;
                }
                Log(0, "UNmountISO result : " + "[" + exitcode + "]" + result);
            }
            catch (Exception ex)
            {
                Log(3, ex.Message);
                ex.GetType();
            }
        }

        public int GetIndexOfDrive(string drive)
        {
            drive = drive.Replace(":", "").Replace(@"\", "");

            // execute DiskPart programatically
            Process process = new Process();
            process.StartInfo.FileName = "diskpart.exe";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.RedirectStandardInput = true;
            process.StartInfo.RedirectStandardOutput = true;
            process.Start();
            process.StandardInput.WriteLine("list disk");
            process.StandardInput.WriteLine("exit");
            string output = process.StandardOutput.ReadToEnd();
            process.WaitForExit();

            // extract information from output
            string table = output.Split(new string[] { "DISKPART>" }, StringSplitOptions.None)[1];
            var rows = table.Split(new string[] { "\n" }, StringSplitOptions.None);
            for (int i = 3; i < rows.Length; i++)
            {
                if (rows[i].Contains("Volume"))
                {
                    int index = Int32.Parse(rows[i].Split(new string[] { " " }, StringSplitOptions.None)[3]);
                    string label = rows[i].Split(new string[] { " " }, StringSplitOptions.None)[8];

                    if (label.Equals(drive))
                    {
                        return index;
                    }
                }
            }

            return -1;
        }

        public int DiskPart(string index,string drive)
        {
            int result = 0;
            try
            {
                string scriptFilePath =@"C:\dpScript.txt";

                if (File.Exists(scriptFilePath))
                {
                    // Note: if the file doesn't exist, no exception is thrown    
                    File.Delete(scriptFilePath); // Delete the script, if it exists    
                }

                // Create the script to resize F and make U    
                string size = (trackBar1.Value * 1024).ToString();

                if (trackBar1.Value > GetTotalSize(drive + "\\"))
                {
                    File.AppendAllText(scriptFilePath,
                    string.Format(
                        "SELECT DISK " + index + "\r\n" + // Select the first disk drive    
                        "CLEAN\r\n" + // Select the drive    
                        "CREATE PARTITION PRIMARY SIZE=" + size + "\r\n" + // Shrink to half the original size    
                        "ASSIGN\r\n" + // Make the drive partition    
                        "ACTIVE\r\n" + // Assign it's letter    
                        "FORMAT FS=fat32 QUICK LABEL=\"BOOT\"\r\n" + // Format it   
                        "CREATE PARTITION PRIMARY\r\n" +
                        "ASSIGN\r\n" +
                        "FORMAT FS=ntfs QUICK LABEL=\"DEPLOY\"\r\n" +
                        "EXIT")
                    ); // And exit    

                    int exitcode = 0;
                    string resultSen = ExecuteCmdCommand("DiskPart.exe" + " /s \"" + scriptFilePath + "\"", ref exitcode);
                    File.Delete(scriptFilePath); // Delete the script file    
                    if (exitcode > 0)
                    {
                        result = exitcode;
                    }
                } else
                {
                    MessageBox.Show("La taille du périphérique USB n'est pas suffisante.", "Périphérique non compatible", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    result = 1;
                }
            } catch(Exception ex)
            {
                Log(3, ex.Message);
                result = 1;
                ex.GetType();
            }
            return result;
       }

        public int GetTotalSize(string drive)
        {
            foreach (DriveInfo d in DriveInfo.GetDrives())
            {
                if (d.IsReady && d.Name == drive)
                {
                    return (int)(d.TotalSize / (1024 * 1024 * 1024));
                }
            }
            return -1;
        }

        private static string ExecuteCmdCommand(string Command, ref int ExitCode)
        {
            ProcessStartInfo ProcessInfo;
            Process Process = new Process();
            string myString = string.Empty;
            ProcessInfo = new ProcessStartInfo("cmd.exe", "/C " + Command);

            ProcessInfo.CreateNoWindow = true;
            ProcessInfo.WindowStyle = ProcessWindowStyle.Hidden;
            ProcessInfo.UseShellExecute = false;
            ProcessInfo.RedirectStandardOutput = true;
            Process.StartInfo = ProcessInfo;
            Process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            Application.DoEvents();
            Process = Process.Start(ProcessInfo);
            Application.DoEvents();
            StreamReader myStreamReader = Process.StandardOutput;
            string line = myStreamReader.ReadLine();
            myString = myStreamReader.ReadToEnd();

            ExitCode = Process.ExitCode;
            Process.Close();

            return myString;
        }

        public static string ExecutePSCommand(string Command, ref int ExitCode)
        {
            ProcessStartInfo ProcessInfo;
            Process Process = new Process();
            string myString = string.Empty;
            ProcessInfo = new ProcessStartInfo("cmd.exe", "/C " + Command);

            ProcessInfo.CreateNoWindow = true;
            ProcessInfo.WindowStyle = ProcessWindowStyle.Hidden;
            ProcessInfo.UseShellExecute = false;
            ProcessInfo.RedirectStandardOutput = true;
            Process.StartInfo = ProcessInfo;
            Process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            Application.DoEvents();
            Process = Process.Start(ProcessInfo);
            Application.DoEvents();
            StreamReader myStreamReader = Process.StandardOutput;
            myString = myStreamReader.ReadLine();
            myStreamReader.ReadToEnd();

            ExitCode = Process.ExitCode;
            Process.Close();

            return myString;
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != "" && comboBox3.Text != "")
            {
                Log(0, "---DEBUT DE LA GENERATION DU DISQUE---");
                DateTime start = DateTime.Now;
                Log(0, "Partition de BOOT séléectionnée : " + comboBox1.Text);
                Log(0, "Chemin des masters :" + textBox1.Text);
                Log(0, "Master sélectionné : " + comboBox3.Text);
                button2.Enabled = false;
                button2.BackColor = Color.FromArgb(224, 222, 216);
                comboBox1.Enabled = false;
                comboBox3.Enabled = false;
                this.Cursor = Cursors.WaitCursor;
                await Task.Run(() =>
                  {
                      timer1.Start();
                      timer1.Interval = 1000; // cad 1 sec
                 });
                Application.DoEvents();
                label3.Text = "Détection de l'espace libre sur C: ...";
                Log(0, label3.Text);
                progressBar1.Value = 5;
                lastVal = progressBar1.Value;
                Application.DoEvents();
                long CSize = GetTotalFreeSpace("C:\\");
                Log(0, "Espace libre sur C:\\ = " + CSize.ToString() + " Go");
                Application.DoEvents();
                string sourceDir = "";
                long CDSize = 0;
                if (IsWindows10())
                {
                    label3.Text = "Montage de l'ISO en lecteur virtuel ...";
                    Log(0, label3.Text);
                    progressBar1.Value = 10;
                    lastVal = progressBar1.Value;
                    Application.DoEvents();
                    string path = Path.Combine(textBox1.Text,comboBox3.Text);
                    string CDROM = await Task.Run(() =>
                    {
                        Application.DoEvents();
                        return MountISO(path);
                    });
                    Application.DoEvents();
                    label3.Text = "Lecteur " + CDROM + ": monté ...";
                    Log(0, label3.Text);
                    progressBar1.Value = 15;
                    lastVal = progressBar1.Value;
                    Application.DoEvents();
                    CDSize = await Task.Run(() =>
                    {
                        Application.DoEvents();
                        return GetTotalFreeSpace(CDROM + ":\\");
                    });
                    Application.DoEvents();
                    sourceDir = CDROM;
                }
                else if (IsWindows7())
                {
                    label3.Text = "Montage de l'ISO en lecteur virtuel ...";
                    Log(0, label3.Text);
                    progressBar1.Value = 10;
                    lastVal = progressBar1.Value;
                    Application.DoEvents();
                    string path = Path.Combine(textBox1.Text, comboBox3.Text);
                    string CDROM = await Task.Run(() =>
                    {
                        Application.DoEvents();
                        return MountISO7(path);
                    });
                    Application.DoEvents();
                    if (CDROM != "" && CDROM != null)
                    {
                        label3.Text = "Lecteur " + CDROM + ": monté ...";
                        Log(0, label3.Text);
                        progressBar1.Value = 15;
                        lastVal = progressBar1.Value;
                        Application.DoEvents();
                        CDSize = await Task.Run(() =>
                        {
                            Application.DoEvents();
                            return GetTotalFreeSpace(CDROM + ":\\");
                        });
                        Application.DoEvents();
                        sourceDir = CDROM;
                    }
                    else
                    {
                        MessageBox.Show("Identification du lecteur virtuel impossible.", "Erreur du lecteur virteul", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        Log(4, "Impossible d'identifier le lecteur virtuel");
                        Application.Exit();
                    }
                }
                else
                {

                }
                if (sourceDir != "")
                {
                    label3.Text = "Vérification de la possibilité de copier les fichiers de l'ISO ...";
                    Log(0, label3.Text);
                    progressBar1.Value = 20;
                    lastVal = progressBar1.Value;
                    Application.DoEvents();
                    if (CSize > CDSize)
                    {
                        Application.DoEvents();
                        label3.Text = "Recherche de la seconde partition ...";
                        Log(0, label3.Text);
                        progressBar1.Value = 25;
                        lastVal = progressBar1.Value;
                        Application.DoEvents();
                        string letter = comboBox1.Text;
                        string partition2 = await Task.Run(() =>
                        {
                            Application.DoEvents();
                            return OtherLetterPartition(letter);
                        });
                        Log(0, "Seconde partition identifiée sur " + partition2);
                        if (partition2 == null)
                        {
                            Application.DoEvents();
                            label3.Text = "En attente ...";
                            Application.DoEvents();
                            MessageBox.Show("La seconde partition ne peut pas être identifiée.\r\nNous ne pouvons donc pas continuer.", "Erreur d'identification", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            Log(3, "Impossible de trouver la seconde partition");
                        }
                        else
                        {
                            Application.DoEvents();
                            label3.Text = "Copie en cours sur " + comboBox1.Text + " des fichiers de Boot ...";
                            Log(0, label3.Text);
                            progressBar1.Value = 30;
                            lastVal = progressBar1.Value;
                            Application.DoEvents();
                            string destDir = comboBox1.Text;
                            CopyBootFiles(sourceDir, destDir);
                            Application.DoEvents();
                            Thread.Sleep(100);
                            Application.DoEvents();
                            //label3.Text = "Copie en cours sur " + partition2 + " du dossier Deploy ...";
                            Log(0, label3.Text);
                            progressBar1.Value = 50;
                            lastVal = progressBar1.Value;
                            Application.DoEvents();
                            CopyDeployDir(sourceDir, partition2);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Impossible d'identifié le lecteur virtuel.", "Erreur lecteur virtuel", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    Log(3, "Impossible d'identifier le lecteur virtuel");
                }
                Application.DoEvents();
                label3.Text = "Ejection du lecteur virtuel ...";
                Log(0, label3.Text);
                progressBar1.Value = 90;
                lastVal = progressBar1.Value;
                Application.DoEvents();
                Thread.Sleep(100);
                Application.DoEvents();
                if (IsWindows10())
                {
                    string path = Path.Combine(textBox1.Text,comboBox3.Text);
                    await Task.Run(() =>
                     {
                         Application.DoEvents();
                         UnMountISO(path);
                     });
                }
                else
                {
                    string path = textBox1.Text.Substring(0, 2);
                    await Task.Run(() =>
                     {
                         Application.DoEvents();
                         UnMountISO7(path);
                     });
                }
                progressBar1.Value = 95;
                lastVal = progressBar1.Value;
                Application.DoEvents();
                label3.Text = "Finalisation ...";
                Log(0, label3.Text);
                Application.DoEvents();
                this.Cursor = Cursors.Default;
                Thread.Sleep(2000);
                timer1.Stop();
                progressBar1.Value = 100;
                Application.DoEvents();
                Log(0, "Opération terminée avec succès.");
                DateTime end = DateTime.Now;
                TimeSpan span = end - start;
                progressBar1.Value = 0;
                Thread.Sleep(5000);
                Log(0, "---FIN DE LA GENERATION DU DISQUE EN : " + String.Format("{0} jour(s), {1}:{2}:{3}", span.Days, span.Hours, span.Minutes, span.Seconds));
                //Process[] processes = (from prc in Process.GetProcessesByName("USBCreator")
                //                       orderby prc.StartTime
                //                       select prc).ToArray();
                //while (processes.Count() != 0)
                //{
                //    Log(0, "Il reste des processes en cours");
                //}
                await Task.Run(() =>
                 {
                     MessageBox.Show("Préparation du disque terminée.", "Fin d'exécution", MessageBoxButtons.OK, MessageBoxIcon.Information);
                 });
            }
            else
            {
                MessageBox.Show("Vous devez choisir : \r\n\t- une partition de Boot\r\n\t- et une image ISO.", "Erreur d'identification", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                Log(2, "Choix de partition de BOOT et/ou de master incorrect");
            }
            button2.Enabled = true;
            button2.BackColor = Color.FromArgb(0, 136, 206);
            comboBox1.Enabled = true;
            comboBox3.Enabled = true;
        }

        public long GetDirectorySize(string fullDirectoryPath)
        {
            long startDirectorySize = 0;
            if (!Directory.Exists(fullDirectoryPath))
                return startDirectorySize; //Return 0 while Directory does not exist.

            var currentDirectory = new DirectoryInfo(fullDirectoryPath);
            //Add size of files in the Current Directory to main size.
            currentDirectory.GetFiles().ToList().ForEach(f => startDirectorySize += f.Length);

            //Loop on Sub Direcotries in the Current Directory and Calculate it's files size.
            currentDirectory.GetDirectories().ToList()
                .ForEach(d => startDirectorySize += GetDirectorySize(d.FullName));

            return startDirectorySize / (1024 * 1024 * 1024);  //Return full Size of this Directory.
        }

        public async Task<string> FormatPartition(string destDirName)
        {
            string result = null;
            try
            {
                string scriptFilePath = @"C:\psScript.ps1";

                if (File.Exists(scriptFilePath))
                {
                    // Note: if the file doesn't exist, no exception is thrown    
                    File.Delete(scriptFilePath); // Delete the script, if it exists    
                }

                // Create the script to resize F and make U    
                File.AppendAllText(scriptFilePath,
                string.Format(
                    "Format-Volume -DriveLetter " + destDirName.First().ToString() + " -FileSystem FAT32 -NewFileSystemLabel BOOT -confirm:$false"
                    )
                ); // And exit    

                int exitcode = 0;
                result = await Task.Run(() =>
                {
                    return ExecutePSCommand("Powershell.exe" + " -executionpolicy bypass -File \"" + scriptFilePath + "\"", ref exitcode);
                });
                File.Delete(scriptFilePath); // Delete the script file    
                if (exitcode > 0)
                {
                    result = null;
                }

                return result;
            }
            catch (Exception ex)
            {
                Log(3, ex.Message);
                return null;
                ex.GetType();
            }
        }

        public async void CopyBootFiles(string CDROM, string destination)
        {
            Application.DoEvents();
            label3.Text = "Nettoyage de " + destination + " ...";
            Log(0, label3.Text);
            Application.DoEvents();
            await FormatPartition(destination);
            Application.DoEvents();

            label3.Text = "Copie du dossier Boot sur " + destination + " ...";
            Log(0, label3.Text);
            Application.DoEvents();
            string SourcePath = CDROM + ":\\Boot";
            string DestinationPath = destination + "\\Boot";
            Application.DoEvents();
            if (!Directory.Exists(DestinationPath))
            {
                Directory.CreateDirectory(DestinationPath);
            }
            Task.Run(() => {
                try
                {
                    FileSystem.CopyDirectory(SourcePath, DestinationPath, UIOption.AllDialogs);
                } catch(Exception ex)
                {
                    Log(3, "ERREUR NON GEREE " + ex.Message);
                    ex.GetType();
                }
            }).Wait();
            Application.DoEvents();

            SourcePath = CDROM + ":\\EFI";
            DestinationPath = destination + "\\EFI";
            Application.DoEvents();
            label3.Text = "Copie du dossier EFI sur " + destination + " ...";
            Log(0, label3.Text);
            Application.DoEvents();
            if (!Directory.Exists(DestinationPath))
            {
                Directory.CreateDirectory(DestinationPath);
            }
            Task.Run(() => {
                try
                {
                    FileSystem.CopyDirectory(SourcePath, DestinationPath, UIOption.AllDialogs);
                }
                catch (Exception ex)
                {
                    Log(3, "ERREUR NON GEREE " + ex.Message);
                    ex.GetType();
                }
            }).Wait();
            Application.DoEvents();

            if (Directory.Exists(CDROM + ":\\sources"))
            {
                SourcePath = CDROM + ":\\sources";
                DestinationPath = destination + "\\sources";
                Application.DoEvents();
                label3.Text = "Copie du dossier sources sur " + destination + " ...";
                Log(0, label3.Text);
                Application.DoEvents();
                if (!Directory.Exists(DestinationPath))
                {
                    Directory.CreateDirectory(DestinationPath);
                }
                Task.Run(() => {
                    try
                    {
                        FileSystem.CopyDirectory(SourcePath, DestinationPath, UIOption.AllDialogs);
                    }
                    catch (Exception ex)
                    {
                        Log(3, "ERREUR NON GEREE " + ex.Message);
                        ex.GetType();
                    }
                }).Wait();
                Application.DoEvents();
            }

            if (Directory.Exists(CDROM + ":\\Deploy\\Boot"))
            {
                SourcePath = CDROM + ":\\Deploy\\Boot";
                DestinationPath = destination + "\\Deploy\\Boot";
                Application.DoEvents();
                label3.Text = "Copie du dossier Deploy\\Boot sur " + destination + " ...";
                Log(0, label3.Text);
                Application.DoEvents();
                if (!Directory.Exists(DestinationPath))
                {
                    Directory.CreateDirectory(DestinationPath);
                }
                Task.Run(() => {
                    try
                    {
                        FileSystem.CopyDirectory(SourcePath, DestinationPath, UIOption.AllDialogs);
                    }
                    catch (Exception ex)
                    {
                        Log(3, "ERREUR NON GEREE " + ex.Message);
                        ex.GetType();
                    }
                }).Wait();
                Application.DoEvents();
            }

            SourcePath = CDROM + ":\\autorun.inf";
            DestinationPath = destination + "\\autorun.inf";
            Application.DoEvents();
            label3.Text = "Copie du fichier autorun.inf sur " + destination + " ...";
            Log(0, label3.Text);
            Application.DoEvents();
            Task.Run(() => {
                try
                {
                    FileSystem.CopyFile(SourcePath, DestinationPath, UIOption.AllDialogs);
                }catch (Exception ex)
                {
                    Log(3, "Erreur sur la copie du fichier, probablement intercepté par l'anti-virus, " + ex.Message);
                    ex.GetType();
                }
            }).Wait();
            Application.DoEvents();

            SourcePath = CDROM + ":\\bootmgr";
            DestinationPath = destination + "\\bootmgr";
            Application.DoEvents();
            label3.Text = "Copie du fihcier bootmgr sur " + destination + " ...";
            Log(0, label3.Text);
            Application.DoEvents();
            Task.Run(() => { FileSystem.CopyFile(SourcePath, DestinationPath, UIOption.AllDialogs); }).Wait();
            Application.DoEvents();

            SourcePath = CDROM + ":\\bootmgr.efi";
            DestinationPath = destination + "\\bootmgr.efi";
            Application.DoEvents();
            label3.Text = "Copie du fichier bootmgr.efi sur " + destination + " ...";
            Log(0, label3.Text);
            Application.DoEvents();
            Task.Run(() => { FileSystem.CopyFile(SourcePath, DestinationPath, UIOption.AllDialogs); }).Wait();
            Application.DoEvents();

            if (File.Exists(CDROM + ":\\setup.exe"))
            {
                SourcePath = CDROM + ":\\setup.exe";
                DestinationPath = destination + "\\setup.exe";
                Application.DoEvents();
                label3.Text = "Copie du fichier setup.exe sur " + destination + " ...";
                Log(0, label3.Text);
                Application.DoEvents();
                Task.Run(() => { FileSystem.CopyFile(SourcePath, DestinationPath, UIOption.AllDialogs); }).Wait();
                Application.DoEvents();
            }

            Application.DoEvents();
            label3.Text = "Identification du master copié dans le fichier iso_install.txt ...";
            Log(0, label3.Text);
            Application.DoEvents();
            string path = destination + "\\iso_install.txt";
            if (!File.Exists(path))
            {
                // Create a file to write to.
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine("ISO actuellement installée est : " + comboBox3.Text);
                    Log(0, "ISO installée est : " + comboBox3.Text);
                }
            }

        }

        public void CopyDeployDir(string CDROM, string destination)
        {
            Application.DoEvents();
            label3.Text = "Suppression du dossier Deploy sur " + destination + " ...";
            Log(0, label3.Text);
            Application.DoEvents();
            if (Directory.Exists(Path.Combine(destination,"Deploy")))
            {
                Task.Run(() => { FileSystem.DeleteDirectory(Path.Combine(destination, "Deploy"), UIOption.OnlyErrorDialogs, RecycleOption.DeletePermanently); }).Wait();
            }
            Application.DoEvents();

            if (Directory.Exists(CDROM + ":\\Deploy"))
            {
                label3.Text = "Copie du dossier Deploy sur " + destination + " ...";
                Log(0, label3.Text);
                Application.DoEvents();
                string SourcePath = CDROM + ":\\Deploy";
                string DestinationPath = destination + "\\Deploy"; ;
                if (!Directory.Exists(DestinationPath))
                {
                    Directory.CreateDirectory(DestinationPath);
                }
                Application.DoEvents();
                Task.Run(() => {
                    try
                    {
                        FileSystem.CopyDirectory(SourcePath, DestinationPath, UIOption.AllDialogs, UICancelOption.ThrowException);
                    }
                    catch (Exception ex)
                    {
                        Log(3, "Erreur sur la copie d'un fichier, probablement intercepté par l'anti-virus, " + ex.Message + ". Une action de l'utilisateur a annulé ou ignoré la copie d'un fichier.");
                        ex.GetType();
                    }
                }).Wait();
                Application.DoEvents();
            }
        }

        public void ListAllMasters()
        {
            this.comboBox3.Text = "";
            string masterDir = textBox1.Text;
            try
            {
                DirectoryInfo d = new DirectoryInfo(masterDir);//Assuming Test is your 
                Application.DoEvents();
                label3.Text = "Vérifie la présence de " + masterDir + " ...";
                Log(0, label3.Text);
                Application.DoEvents();
                if (Directory.Exists(masterDir))
                {
                    try
                    {
                        this.comboBox3.Enabled = false;
                        this.button2.Enabled = false;
                        this.button2.BackColor = Color.FromArgb(224, 222, 216);
                        this.comboBox3.Items.Clear();
                        FileInfo[] Files = d.GetFiles("*.iso"); //Getting iso files
                        Application.DoEvents();
                        label3.Text = "Remplissage de la liste des masters ...";
                        Log(0, label3.Text);
                        Application.DoEvents();
                        this.comboBox3.Items.Add("");
                        foreach (FileInfo file in Files)
                        {
                            this.comboBox3.Enabled = true;
                            this.button2.Enabled = true;
                            button2.BackColor = Color.FromArgb(0, 136, 206);
                            this.comboBox3.Items.Add(file.Name);
                            Log(0, "Ajout de l'ISO : " + file.Name);
                        }
                        comboBox3.SelectedIndex = 0;
                    }
                    catch (Exception ex)
                    {
                        Log(3, ex.Message);
                        ex.GetType();
                    }
                } else
                {
                    try
                    {
                        textBox1.Text = "";
                    }
                    catch (Exception ex)
                    {
                        Log(3, ex.Message);
                        ex.GetType();
                    }
                }
            }
            catch (Exception ex)
            {
                Log(3, ex.Message);
                ex.GetType();
            }


            Application.DoEvents();
            label3.Text = "En attente ...";
            Application.DoEvents();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            ListAllMasters();
            if (comboBox1.Items.Count <= 1)
            {
                comboBox1.Enabled = false;
                button2.Enabled = false;
                button2.BackColor = Color.FromArgb(224, 222, 216);
            }
            else
            {
                comboBox1.Enabled = true;
            }

            if (comboBox3.Items.Count <= 1)
            {
                comboBox3.Enabled = false;
                button2.Enabled = false;
                button2.BackColor = Color.FromArgb(224, 222, 216);
            }
            else
            {
                comboBox3.Enabled = true;
            }
            if (comboBox1.Enabled == true && comboBox3.Enabled == true)
            {
                button2.Enabled = true;
                button2.BackColor = Color.FromArgb(0, 136, 206);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SetAllUSBDriveFat32();
            if (comboBox1.Items.Count <= 1)
            {
                comboBox1.Enabled = false;
                button2.Enabled = false;
                button2.BackColor = Color.FromArgb(224, 222, 216);
            }
            else
            {
                comboBox1.Enabled = true;
            }
            if (comboBox3.Items.Count <= 1)
            {
                comboBox3.Enabled = false;
                button2.Enabled = false;
                button2.BackColor = Color.FromArgb(224, 222, 216);
            }
            else
            {
                comboBox3.Enabled = true;
            }
            if (comboBox1.Enabled == true && comboBox3.Enabled == true)
            {
                button2.Enabled = true;
                button2.BackColor = Color.FromArgb(0, 136, 206);
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                int val = progressBar1.Value;
                Application.DoEvents();
                val += (val % progressBar1.Maximum) + 1;
                if (val > progressBar1.Maximum)
                {
                    val = lastVal;
                }
                Application.DoEvents();
                progressBar1.Value = val;
                Application.DoEvents();
            } catch(Exception ex)
            {
                ex.GetType();
            }
        }

        private void timer2_Tick_1(object sender, EventArgs e)
        {
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            ListAllMasters();
            Registry.SetValue(HKLMUsbCreator, "MasterPath", textBox1.Text, RegistryValueKind.String);
            Registry.SetValue(HKLMUsbCreator, "lastMasterUsed", comboBox3.Text, RegistryValueKind.String);
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Registry.SetValue(HKLMUsbCreator, "Version", Application.ProductVersion, RegistryValueKind.String);
            Registry.SetValue(HKLMUsbCreator, "MasterPath", textBox1.Text, RegistryValueKind.String);
            Registry.SetValue(HKLMUsbCreator, "lastMasterUsed", comboBox3.Text, RegistryValueKind.String);
            Registry.SetValue(HKLMUsbCreator, "PartSize", trackBar1.Value, RegistryValueKind.String);
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Registry.SetValue(HKLMUsbCreator, "Version", Application.ProductVersion, RegistryValueKind.String);
            Registry.SetValue(HKLMUsbCreator, "MasterPath", textBox1.Text, RegistryValueKind.String);
            Registry.SetValue(HKLMUsbCreator, "lastMasterUsed", comboBox3.Text, RegistryValueKind.String);
            Registry.SetValue(HKLMUsbCreator, "PartSize", trackBar1.Value, RegistryValueKind.String);
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            Registry.SetValue(HKLMUsbCreator, "PartSize", trackBar1.Value, RegistryValueKind.String);
        }

        private void comboBox3_SelectionChangeCommitted(object sender, EventArgs e)
        {
            Registry.SetValue(HKLMUsbCreator, "MasterPath", textBox1.Text, RegistryValueKind.String);
            Registry.SetValue(HKLMUsbCreator, "lastMasterUsed", comboBox3.Text, RegistryValueKind.String);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var reg = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion");
            string OSInstalled = (string)reg.GetValue("ProductName");
            string releaseId = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ReleaseId", "").ToString();
            string version = GetWindowsVersion();
            string supportInfos = "Bonjour,%0D%0AUSBCREATOR : " + Application.ProductVersion + "%0D%0AOS : " + OSInstalled + "%0D%0ABuild : " + releaseId + "%0D%0ARévision : " + version + "%0D%0AVeuillez décrire votre problème ci-dessous : %0D%0A%0D%0A";
            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            proc.StartInfo.FileName = "mailto:masters@sncf.fr?subject=[USBCreator]Support&body=" + supportInfos;
            proc.Start();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.services.dsit.sncf.fr/mastersncf/");
        }

        public string GetWindowsVersion()
        {
            var reg = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion");
            string revision = (string)reg.GetValue("UBR").ToString();
            string currentbuild = (string)reg.GetValue("CurrentBuild").ToString();
            return currentbuild + "." + revision;
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
