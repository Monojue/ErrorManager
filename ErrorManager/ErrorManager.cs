using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ErrorManager {
    public partial class ErrorManager : Form {
        public ErrorManager() {
            InitializeComponent();
            lblCondition.Text = "";
        }

        List<string> files;

        private void btnStart_Click(object sender, EventArgs e) {
            if(!backgroundWorker.IsBusy && !tbErrorPath.Text.Equals(string.Empty)){
                lblCondition.Text = "";
                backgroundWorker.RunWorkerAsync();
            }
        }

        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e) {
            files = Directory.GetFiles(tbErrorPath.Text).ToList();
            ErrorListBook book = new ErrorListBook();
            book.GetErrorListFormExcel(files, tbErrorPath.Text, sender as BackgroundWorker);
        }

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e) {
            lblProgress.Text = e.ProgressPercentage.ToString() + "/" + files.Count.ToString();
        }

        private void btnCancel_Click(object sender, EventArgs e) {
            backgroundWorker.CancelAsync();
            lblCondition.Text = "User Cancel!";
        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
            if(lblCondition.Text.Equals("")){
                lblCondition.Text = "File Exported to your Path Successfully!";
            }
        }

        private void tbErrorPath_MouseClick(object sender, MouseEventArgs e) {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog()){
                fbd.SelectedPath = @"C:\UnitTest\S";
                if (fbd.ShowDialog() == DialogResult.OK) {
                    tbErrorPath.Text = fbd.SelectedPath;
                }
            }
        }
    }
}
