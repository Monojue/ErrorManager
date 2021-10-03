namespace ErrorManager {
    partial class ErrorManager {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.tbErrorPath = new System.Windows.Forms.TextBox();
            this.btnStart = new System.Windows.Forms.Button();
            this.backgroundWorker = new System.ComponentModel.BackgroundWorker();
            this.lblProgress = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.lblCondition = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // tbErrorPath
            // 
            this.tbErrorPath.Location = new System.Drawing.Point(12, 12);
            this.tbErrorPath.Name = "tbErrorPath";
            this.tbErrorPath.Size = new System.Drawing.Size(370, 19);
            this.tbErrorPath.TabIndex = 0;
            this.tbErrorPath.MouseClick += new System.Windows.Forms.MouseEventHandler(this.tbErrorPath_MouseClick);
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(404, 10);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(75, 23);
            this.btnStart.TabIndex = 1;
            this.btnStart.Text = "Start";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // backgroundWorker
            // 
            this.backgroundWorker.WorkerReportsProgress = true;
            this.backgroundWorker.WorkerSupportsCancellation = true;
            this.backgroundWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker_DoWork);
            this.backgroundWorker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker_ProgressChanged);
            this.backgroundWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker_RunWorkerCompleted);
            // 
            // lblProgress
            // 
            this.lblProgress.AutoSize = true;
            this.lblProgress.Font = new System.Drawing.Font("MS UI Gothic", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblProgress.Location = new System.Drawing.Point(138, 48);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Size = new System.Drawing.Size(14, 13);
            this.lblProgress.TabIndex = 2;
            this.lblProgress.Text = "/";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("MS UI Gothic", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.Location = new System.Drawing.Point(12, 48);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(57, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Progress";
            // 
            // lblCondition
            // 
            this.lblCondition.AutoSize = true;
            this.lblCondition.Font = new System.Drawing.Font("MS UI Gothic", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblCondition.Location = new System.Drawing.Point(247, 75);
            this.lblCondition.Name = "lblCondition";
            this.lblCondition.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblCondition.Size = new System.Drawing.Size(233, 13);
            this.lblCondition.TabIndex = 4;
            this.lblCondition.Text = "File Exported to your Path Successfully!";
            this.lblCondition.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(404, 44);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 5;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // ErrorManager
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(492, 97);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.lblCondition);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblProgress);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.tbErrorPath);
            this.Name = "ErrorManager";
            this.Text = "ErrorManager";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tbErrorPath;
        private System.Windows.Forms.Button btnStart;
        private System.ComponentModel.BackgroundWorker backgroundWorker;
        private System.Windows.Forms.Label lblProgress;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblCondition;
        private System.Windows.Forms.Button btnCancel;
    }
}