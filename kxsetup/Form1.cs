using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using Microsoft.Win32;
using IWshRuntimeLibrary;
using System.Threading;
using System.Runtime.InteropServices;
using Ionic.Zip;
using System.Management;




namespace kxsetup
{

    public partial class Form1 : Form
    {
        string ProgramFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);

        string Temp = Environment.GetEnvironmentVariable("TEMP");
        string Programs = Environment.GetFolderPath(Environment.SpecialFolder.Programs);
        string sys = Environment.GetFolderPath(Environment.SpecialFolder.System);
        string GetCurrentDirectory = Directory.GetCurrentDirectory();//取运行目录
        bool bit;
        int by = 0;
        int jindu = 0;
        int jindulable = 0;
        Form2 f2 = new Form2();
        [DllImport("user32.dll", EntryPoint = "SendMessage")]
        private static extern int SendMessage(IntPtr hwnd, int wMsg, int wParam, int lParam);
        [DllImport("User32.dll", EntryPoint = "FindWindow")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("newdev.dll", EntryPoint = "DiInstallDriver")]
        private static extern bool DiInstallDriver(int hwnd, string path, int d, bool s);
        [DllImport("kernel32.dll")]
        public static extern uint GetLastError();
        [DllImport("user32.dll", EntryPoint = "ReleaseCapture")]
        public static extern bool ReleaseCapture();
        public const int WM_SYSCOMMAND = 0x0112;
        public const int SC_MOVE = 0xF010;
        public const int HTCAPTION = 0x0002;
        public const string MD5 = "7c2195fcf2ae52fd753e6dd13a9c5549";//ui
        public const string MD52 = "7438bf3dfb8ce0ef93f92f9d7b88c2d0";
        public Form1()
        {
            InitializeComponent();
            this.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            byte[] UI = Properties.Resources.UI;
            System.IO.File.WriteAllBytes(Temp + @"\UI.png", UI);
            byte[] guanyu = Properties.Resources.关于;
            System.IO.File.WriteAllBytes(Temp + @"\关于.png", guanyu);

            f2.Show();
            if (GetMD5HashFromFile(Temp + @"\UI.png") != MD5 || GetMD5HashFromFile(Temp + @"\关于.png") != MD52)
            {
                MessageBox.Show("检测到未经授权的修改，请重新下载本程序", "发现盗版", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(0);

            }
            else
            {
                this.BackgroundImage = Image.FromFile(Temp + @"\UI.png");
                f2.BackgroundImage = Image.FromFile(Temp + @"\关于.png");
            }


            init();

        }


        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult dr;
            dr = MessageBox.Show("准备安装，请不要进行其他操作。要开始安装吗？", "开始安装", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
            if (dr == DialogResult.Yes)
            {
                button2.Enabled = false;
                Control.CheckForIllegalCrossThreadCalls = false;
                Thread t = new Thread(new ThreadStart(anzhuang));
                t.IsBackground = true;
                t.Start();
                label2.Visible = true;
            }

        }
        static string GetSystem(string path)
        {
            new FileInfo(path);
            FileVersionInfo info = FileVersionInfo.GetVersionInfo(path);

            return (info.FileVersion);

        }


        private static void CreateDesktopLnk(string cut, string target, string argument, string des, string working, string icon)
        {

            WshShell shell = new WshShellClass();

            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(cut);//输出路径
            shortcut.TargetPath = target;//文件路径
            shortcut.Arguments = argument;// 参数
            shortcut.Description = des;//备注
            shortcut.WorkingDirectory = working;//程序所在文件夹，在快捷方式图标点击右键可以看到此属性
            shortcut.IconLocation = icon;//图标
            shortcut.Hotkey = "";
            shortcut.WindowStyle = 1;
            shortcut.Save();
        }



        private void zipfile(string zipToUnpack, string unpackDirectory)
        {

            using (ZipFile zip1 = ZipFile.Read(zipToUnpack))
            {
                // here, we extract every entry, but we could extract conditionally
                // based on entry name, size, date, checkbox status, etc.  
                foreach (ZipEntry e in zip1)
                {
                    e.Extract(unpackDirectory, ExtractExistingFileAction.OverwriteSilently);
                }
            }
        }
        public static string GetSoundPNPID()
        {
            string st = "";
            ManagementObjectSearcher mos = new ManagementObjectSearcher("Select * from Win32_SoundDevice");
            foreach (ManagementObject mo in mos.Get())
            {
                st = mo["Name"].ToString();
            }
            return st;
        }
        public static void Delay(int milliSecond)
        {
            int start = Environment.TickCount;
            while (Math.Abs(Environment.TickCount - start) < milliSecond)
            {
                Application.DoEvents();
            }
        }

        private void Form1_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0);
        }

        private static string GetMD5HashFromFile(string fileName)
        {
            StringBuilder sb = new StringBuilder();
            try
            {

                FileStream file = new FileStream(fileName, FileMode.Open);
                System.Security.Cryptography.MD5 md5 = new System.Security.Cryptography.MD5CryptoServiceProvider();
                byte[] retVal = md5.ComputeHash(file);
                file.Close();


                for (int i = 0; i < retVal.Length; i++)
                {
                    sb.Append(retVal[i].ToString("x2"));
                }
                return sb.ToString();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return sb.ToString();
                Environment.Exit(0);
            }
        }
        private void jindutiao()
        {
            if (jindu < 475)
            {
                jindu = jindu + 59;
                jindulable = jindulable + 10;
                pictureBox2.Size = new Size(jindu, 26);
                label1.Text = jindulable + "%";
            }
            else if (jindu < 594)
            {
                pictureBox2.Size = new Size(594, 26);
                label1.Text = "100%";
            }



        }
        private void anzhuang()
        {
            string ProgramFilesX86 = ProgramFiles + @" (x86)";
            try
            {
                jindutiao();//1
                Process kxprocess;
                bool dev;
                byte[] x86 = Properties.Resources.x86;
                byte[] x64 = Properties.Resources.x64;
                byte[] reg = Properties.Resources.kx;
                byte[] lzip = Properties.Resources.lzip;
                byte[] llb = Properties.Resources.llb;
                if (System.IO.File.Exists(GetCurrentDirectory + @"\Ionic.Zip.dll") || System.IO.File.Exists(GetCurrentDirectory + @"\Interop.IWshRuntimeLibrary.dll"))
                {
                    System.IO.File.Delete(GetCurrentDirectory + @"\Ionic.Zip.dll");
                    System.IO.File.Delete(GetCurrentDirectory + @"\Interop.IWshRuntimeLibrary.dll");

                }
                System.IO.File.WriteAllBytes(GetCurrentDirectory + @"\Ionic.Zip.dll", lzip);
                System.IO.File.WriteAllBytes(GetCurrentDirectory + @"\Interop.IWshRuntimeLibrary.dll", llb);
                FileInfo fi = new FileInfo(GetCurrentDirectory + @"\Ionic.Zip.dll");
                System.IO.File.SetAttributes(GetCurrentDirectory + @"\Ionic.Zip.dll", fi.Attributes | FileAttributes.Hidden);
                FileInfo fb = new FileInfo(GetCurrentDirectory + @"\Ionic.Zip.dll");
                System.IO.File.SetAttributes(GetCurrentDirectory + @"\Interop.IWshRuntimeLibrary.dll", fb.Attributes | FileAttributes.Hidden);
                System.IO.File.WriteAllBytes(Temp + @"\x86.zip", x86);
                System.IO.File.WriteAllBytes(Temp + @"\x64.zip", x64);
                jindutiao();//2
                zipfile(Temp + @"\x86.zip", ProgramFilesX86);
                jindutiao();//3
                zipfile(Temp + @"\x64.zip", ProgramFiles);
                jindutiao();//4
                System.IO.File.WriteAllBytes(ProgramFiles + @"\kX Project\kx.reg", reg);
                //Infprocess = Process.Start(sys + @"\InfDefaultInstall.exe", "\"" + ProgramFiles + @"\kX Project\kx.inf" + "\"");
                dev = DiInstallDriver(0, ProgramFiles + @"\kX Project\kx.inf", 0, false);
                if (dev == false)
                {
                    MessageBox.Show("设备安装失败，请卸载后再试", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Environment.Exit(0);
                }
                //MessageBox.Show(GetLastError().ToString());返回设备安装状态
                jindutiao();//5
                Thread.Sleep(3000);
                Process.Start("regedit", @"/s" + " \"" + ProgramFiles + @"\kX Project\kx.reg" + "\"");
                jindutiao();//6
                            // Infprocess.Kill();
                            //  Infprocess.Close();

                if (by == 1)
                {
                    Registry.SetValue(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", ProgramFilesX86 + @"\kX Project\kxmixer.exe", "~ WIN7RTM");//设置兼容性

                }


                CreateDesktopLnk(ProgramFiles + @"\kX Project\kxsetup.lnk", ProgramFiles + @"\kX Project\kxsetup.exe", "--reset kxsetup", "图丁网络", ProgramFiles + @"\kX Project\", ProgramFiles + @"\kX Project\kxsetup.exe");
                Process.Start(ProgramFiles + @"\kX Project\kxsetup.lnk");
                jindutiao();//7

                if (!Directory.Exists(Programs + @"\kX Audio Driver\"))
                {
                    Directory.CreateDirectory(Programs + @"\kX Audio Driver\");
                }
                CreateDesktopLnk(Programs + @"\kX Audio Driver\kX Mixer (32-bit).lnk", ProgramFilesX86 + @"\kX Project\kxmixer.exe", "", "图丁网络", ProgramFilesX86 + @"\kX Project\", ProgramFilesX86 + @"\kX Project\kxmixer.exe");
                CreateDesktopLnk(Programs + @"\kX Audio Driver\Uninstall kX Project audio driver.lnk", ProgramFiles + @"\kX Project\Uninstall.exe", "", "图丁网络", ProgramFiles + @"\kX Project\", ProgramFiles + @"\kX Project\Uninstall.exe");
                jindutiao();//8

                kxprocess = Process.Start(Programs + @"\kX Audio Driver\kX Mixer (32-bit).lnk");
                Thread.Sleep(1000);
                kxprocess.Kill();
                kxprocess.Close();
                Thread.Sleep(1000);
                Process.Start(Programs + @"\kX Audio Driver\kX Mixer (32-bit).lnk");
                jindutiao();//9
                Thread.Sleep(1000);
                string st = GetSoundPNPID();
                RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run\", true);
                key.DeleteValue("kX Mixer", true);
                key.SetValue("kX Mixer", "\"" + ProgramFilesX86 + @"\kX Project\kxmixer.exe" + "\"" + " --startup");
                key.Close();
                jindutiao();//10
                button1.BackgroundImage = Properties.Resources.安装完成灰;
                button1.Enabled = false;
                label2.Visible = false;
                init();
                //MessageBox.Show(st);
                if (st == "kX 10k1 Audio (3552) - Generic")
                {
                    Process.Start(ProgramFilesX86 + @"\kX Project\5.1.kx");
                    Process.Start("rundll32.exe", "shell32.dll,Control_RunDLL mmsys.cpl @1");

                    MessageBox.Show("请将：\n1.播放设备：SPDIF/AC3 Output设为默认值\n2.录制设备：Rcording Mixer设为默认值", "安装完成", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                else if (st == "kX 10k2 Audio (3552) - Generic")
                {
                    Process.Start(ProgramFilesX86 + @"\kX Project\7.1.kx");
                    Process.Start("rundll32.exe", "shell32.dll,Control_RunDLL mmsys.cpl @1");

                    MessageBox.Show("请将：\n1.播放设备：SPDIF/AC3 Output设为默认值\n2.录制设备：Rcording Mixer设为默认值", "安装完成", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                else if (st == "kX 10k2.5 Audio (3552) - Generic")
                {
                    Process.Start(ProgramFilesX86 + @"\kX Project\7.1.kx");
                    Process.Start("rundll32.exe", "shell32.dll,Control_RunDLL mmsys.cpl @1");

                    MessageBox.Show("请将：\n1.播放设备：SPDIF/AC3 Output设为默认值\n2.录制设备：Rcording Mixer设为默认值", "安装完成", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                else
                {
                    MessageBox.Show("1.可能设备安装失败，请卸载后重新安装。\n2.可能您的计算机中不存在本驱动支持的声卡设备\n3.可能您的计算机中有多块声卡我们无法识别您的声卡类型，请手动选择效果文件并将播放设备：SPDIF/AC3 Output设为默认值 录制设备：Rcording Mixer设为默认值", "未识别的声卡类型", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }


                //Registry.SetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "kX Mixer", "\"" + ProgramFilesX86 + @"\kX Project\kxmixer.exe" + "\"" + " --startup");//设置开机启动项
                button2.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                button2.Enabled = true;
            }

        }
        private void init()//初始化
        {

            if (IntPtr.Size == 8)
            {
                bit = true;
            }
            else if (IntPtr.Size == 4)
            {
                bit = false;
            }
            else
            {
                //...NotSupport
            }
            pictureBox2.Size = new Size(0, 26);
            label1.Text = "";

            string i = GetSystem(sys + @"\kernel32.dll");
            string[] s = i.Split(new char[] { '.' });
            string ss = s[0] + "." + s[1];
            String sysinfo;
            String bitinfo;
            string dev = "检测中";
            if (double.Parse(ss) >= 10.0)
            {
                by = 1;

            }

            if (bit == false || double.Parse(ss) < 6.1)
            {
                MessageBox.Show("抱歉目前仅可在Windows7及以上64位系统中安装", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }

            if (double.Parse(ss) == 5.1)
            {
                sysinfo = "Windows XP";

            }
            else if (double.Parse(ss) == 6.0)
            {
                sysinfo = "Windows Vista";

            }
            else if (double.Parse(ss) == 6.1)
            {

                sysinfo = "Windwos 7";
            }
            else if (double.Parse(ss) == 6.2)
            {
                sysinfo = "Windows 8";

            }
            else if (double.Parse(ss) == 6.3)
            {
                sysinfo = "Windows 8.1";
            }
            else if (double.Parse(ss) == 10.0)
            {
                sysinfo = "Windows 10";
            }
            else
            {
                sysinfo = "Windows NT" + ss;
            }
            if (bit = true)
            {
                bitinfo = "64";
            }
            else
            {
                bitinfo = "32";
            }
            if (System.IO.File.Exists(ProgramFiles + @"\kX Project\kxmixer.exe") != true)
            {
                dev = "未安装";

            }
            else if (Process.GetProcessesByName("kxmixer").Length < 1)

            {

                dev = "已安装，但未运行";
                button1.BackgroundImage = Properties.Resources.开始安装灰;
                button1.Enabled = false;
                label2.Visible = true;
                label2.Text = "无法开始安装，请卸载KX驱动文件重启后再试";
            }
            else
            {
                dev = "已安装";
                button1.BackgroundImage = Properties.Resources.开始安装灰;
                button1.Enabled = false;
                label2.Visible = true;
                label2.Text = "感谢您使用本程序，您的KX驱动可以正常使用，无需安装";
            }

            label3.Text = "系统版本：" + sysinfo + "\n" + "系统位数：" + bitinfo + "\n" + "驱动版本：3552\n" + "软件版本：1.0\n" + "驱动状态：" + dev;
        }
        //事件
        private void button2_Click(object sender, EventArgs e)
        {


            this.Close();

        }

        private void button2_MouseEnter(object sender, EventArgs e)
        {
            button2.BackgroundImage = Properties.Resources.激活叉;
        }

        private void button2_MouseLeave(object sender, EventArgs e)
        {
            button2.BackgroundImage = Properties.Resources.叉;
        }

        private void button1_MouseDown(object sender, MouseEventArgs e)
        {
            button1.BackgroundImage = Properties.Resources.开始安装点击;
        }

        private void button1_MouseUp(object sender, MouseEventArgs e)
        {
            button1.BackgroundImage = Properties.Resources.开始安装绿;
        }

        private void button2_MouseDown(object sender, MouseEventArgs e)
        {
            button2.BackgroundImage = Properties.Resources.叉点击;
        }

        private void button3_MouseDown(object sender, MouseEventArgs e)
        {
            button3.BackgroundImage = Properties.Resources.最小化点击;
        }

        private void button3_MouseEnter(object sender, EventArgs e)
        {
            button3.BackgroundImage = Properties.Resources.激活最小化;
        }

        private void button3_MouseLeave(object sender, EventArgs e)
        {
            button3.BackgroundImage = Properties.Resources.最小化;
        }

        private void button4_MouseEnter(object sender, EventArgs e)
        {
            button4.BackgroundImage = Properties.Resources.激活菜单;
        }

        private void button4_MouseLeave(object sender, EventArgs e)
        {
            button4.BackgroundImage = Properties.Resources.菜单;
        }

        private void button4_MouseDown(object sender, MouseEventArgs e)
        {
            button4.BackgroundImage = Properties.Resources.菜单点击;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }



        private void button4_Click(object sender, EventArgs e)
        {
            contextMenuStrip1.Show(button4, new Point(0, button4.Height + 5));
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (System.IO.File.Exists(GetCurrentDirectory + @"\del.bat"))
            {
                System.IO.File.Delete(GetCurrentDirectory + @"\del.bat");
            }
            if (System.IO.File.Exists(GetCurrentDirectory + @"\Ionic.Zip.dll") || System.IO.File.Exists(GetCurrentDirectory + @"\Interop.IWshRuntimeLibrary.dll"))
            {

                System.IO.File.WriteAllBytes(GetCurrentDirectory + @"\del.bat", Properties.Resources.del);
                FileInfo fc = new FileInfo(GetCurrentDirectory + @"\Ionic.Zip.dll");
                System.IO.File.SetAttributes(GetCurrentDirectory + @"\del.bat", fc.Attributes | FileAttributes.Hidden);
                Process proc = new Process();
                proc.StartInfo.FileName = GetCurrentDirectory + @"\del.bat";
                proc.StartInfo.CreateNoWindow = true;
                proc.StartInfo.UseShellExecute = false;
                proc.Start();
            }



            System.Environment.Exit(0);
        }

        private void 关于我们ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            f2.Opacity = 100;
            f2.ShowInTaskbar = true;
        }
    }
}

