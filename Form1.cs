using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace FileBackupZipper
{
    public partial class BankFileZipper : Form
    {
        const string SC2_PROCESS_NAME = "sc2";

        public BankFileZipper()
        {
            InitializeComponent();

           // label3.Text = Properties.Settings.Default["BankFilePath"].ToString();
           // label4.Text = Properties.Settings.Default["OutputFolder"].ToString();
            label7.Text = (Properties.Settings.Default.LastBackupDate != default(DateTime)) ? Properties.Settings.Default.LastBackupDate.ToString("yyyy-MM-dd HH:mm") : "Never";

            pictureBox1.BackColor = System.Drawing.Color.Red;
            ListenToProgramStartup();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            string defaultFolder = "";
            if (label3.Text != "")
            {
                defaultFolder = label3.Text;
            }
            
            string path = SelectFolder(defaultFolder);
            if (path != "")
            {
                label3.Text = path;

                Properties.Settings.Default["BankFilePath"] = path;
                Properties.Settings.Default.Save();
            };
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string defaultFolder = "";
            if (label4.Text != "") {
                defaultFolder = label4.Text;
            }
            
            string path = SelectFolder(defaultFolder);
            if (path != "") {
                label4.Text = path;
                Properties.Settings.Default["OutputFolder"] = path;
                Properties.Settings.Default.Save();
            };
        }

        private string SelectFolder(string defaultFolder)
        {
            if (defaultFolder == "")
            {
                defaultFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }

            CommonOpenFileDialog dialog = new CommonOpenFileDialog
            {
                InitialDirectory = defaultFolder,
                IsFolderPicker = true
            };

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
               return dialog.FileName;
            }

            return "";
        }

        private void ZipIt(string startPath, string outputPath)
        {
             if(!Directory.Exists(startPath) ||
                !Directory.Exists(outputPath)
             ){
                MessageBox.Show(String.Format("BankfileLocation folder and/or Output folder are not set"), "Error");
                return;
            }

            try
            {
                ZipFile.CreateFromDirectory(startPath, outputPath, CompressionLevel.Fastest, true);

                Properties.Settings.Default["LastBackupDate"] = DateTime.Now;
                Properties.Settings.Default.Save();

                label7.Invoke((MethodInvoker)delegate
                {
                    label7.Text = Properties.Settings.Default.LastBackupDate.ToString("yyyy-MM-dd HH:mm");
                });
            }
            catch (Exception e) {
                MessageBox.Show(String.Format(e.Message), "Error");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ZipIt(@label3.Text, @label4.Text + "/backup-" + DateTime.Now.ToString("yyyyMMddHHmm") + ".zip");
        }

        private async void ListenToProgramStartup()
        {
            var searching = true;

            while (searching) {
                await Task.Run(async () => {
                    searching = !(await IsProcessRunning(SC2_PROCESS_NAME));
                    await Task.Delay(2000);
                });
            }

            ListenToProgramStartup();
        }

        private void SetProcessFound()
        {
            label5.Invoke((MethodInvoker)delegate {
                label5.Text = "Starcraft found!";
            });

            pictureBox1.Invoke((MethodInvoker)delegate{
                pictureBox1.BackColor = System.Drawing.Color.Green;
            });
        }

        private void SetProcessLost()
        {
            label5.Invoke((MethodInvoker)delegate {
                label5.Text = "Starcraft not running";
            });

            pictureBox1.Invoke((MethodInvoker)delegate {
                pictureBox1.BackColor = System.Drawing.Color.Red;
            });
        }

        private async Task<bool> IsProcessRunning(string sProcessName)
        {
            Process[] proc = Process.GetProcessesByName(sProcessName);
            if (proc.Length > 0)
            {
                SetProcessFound();

                await Task.Run(() => {
                    proc[0].WaitForExit();
                    ZipIt(@label3.Text, @label4.Text + "/backup-"+ DateTime.Now.ToString("yyyyMMddHHmm") + ".zip");
                });

                return true;
            }

            SetProcessLost();

            return false;
        }

        private void label3_Click(object sender, EventArgs e)
        {
            OpenExplorer(((Label)sender).Text);
        }

        private void label4_Click(object sender, EventArgs e)
        {
            OpenExplorer(((Label)sender).Text);
        }

        private void OpenExplorer(string path)
        {
            if (path == "Empty")
            {
                return;
            }

            Process.Start("explorer.exe", "\"" + path + "\"");
        }
    }
}
