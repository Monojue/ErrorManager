using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ErrorManager {
    public partial class ErrorManager : Form {
        public ErrorManager() {
            InitializeComponent();
            lblProgress.Text = "";
        }

        List<string> files;

        private void btnStart_Click(object sender, EventArgs e) {
            if(!backgroundWorker.IsBusy && !tbErrorPath.Text.Equals(string.Empty)){
                lblProgress.Text = "Starting...";
                backgroundWorker.RunWorkerAsync();
            }
        }

        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e) {
            Regex reg = new Regex(@"(\d+_\w+_\d+_Error)");
            files = Directory.GetFiles(tbErrorPath.Text).Where(path => reg.IsMatch(path)).ToList();
            ErrorListBook book = new ErrorListBook();
            Boolean cancel = book.GetErrorListFormExcel(files, tbErrorPath.Text, sender as BackgroundWorker);
            if(cancel) {
                e.Cancel = true;
            }
        }

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e) {
            lblProgress.Text = e.ProgressPercentage.ToString() + "/" + files.Count.ToString();
        }

        private void btnCancel_Click(object sender, EventArgs e) {
            if (backgroundWorker.WorkerSupportsCancellation == true) {
                // Cancel the asynchronous operation.
                backgroundWorker.CancelAsync();
            }
        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
           
            if (e.Cancelled == true) {
                lblProgress.Text = "Canceled!";
            } else if (e.Error != null) {
                lblProgress.Text = "Error: " + e.Error.Message;
            } else {
                lblProgress.Text = "File Exported to your Path Successfully!";
            }
        }

        private void tbErrorPath_MouseClick(object sender, MouseEventArgs e) {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog()){
                fbd.SelectedPath = @"C:\UnitTest\S";
                if (fbd.ShowDialog(this) == DialogResult.OK) {
                    tbErrorPath.Text = fbd.SelectedPath;
                }
            }
        }
    }
}
