namespace Peppermint_Outlook_AddIn
{
    partial class frmRecordAudio
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnAttachAudio = new System.Windows.Forms.Button();
            this.lblRecordTimer = new System.Windows.Forms.Label();
            this.lblStop = new System.Windows.Forms.Label();
            this.txtMessage = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.PlayButton = new System.Windows.Forms.PictureBox();
            this.PauseButton = new System.Windows.Forms.PictureBox();
            this.ProgressBar = new System.Windows.Forms.ProgressBar();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.PlayButton)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.PauseButton)).BeginInit();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.ForeColor = System.Drawing.Color.White;
            this.btnCancel.Location = new System.Drawing.Point(6, 202);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(169, 32);
            this.btnCancel.TabIndex = 0;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnAttachAudio
            // 
            this.btnAttachAudio.BackColor = System.Drawing.Color.LightSeaGreen;
            this.btnAttachAudio.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnAttachAudio.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAttachAudio.ForeColor = System.Drawing.SystemColors.Window;
            this.btnAttachAudio.Location = new System.Drawing.Point(179, 202);
            this.btnAttachAudio.Name = "btnAttachAudio";
            this.btnAttachAudio.Size = new System.Drawing.Size(169, 32);
            this.btnAttachAudio.TabIndex = 1;
            this.btnAttachAudio.Text = "Attach";
            this.btnAttachAudio.UseVisualStyleBackColor = false;
            this.btnAttachAudio.Click += new System.EventHandler(this.btnAttachAudio_Click);
            // 
            // lblRecordTimer
            // 
            this.lblRecordTimer.AutoSize = true;
            this.lblRecordTimer.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRecordTimer.ForeColor = System.Drawing.Color.DarkGray;
            this.lblRecordTimer.Location = new System.Drawing.Point(156, 88);
            this.lblRecordTimer.Name = "lblRecordTimer";
            this.lblRecordTimer.Size = new System.Drawing.Size(44, 16);
            this.lblRecordTimer.TabIndex = 6;
            this.lblRecordTimer.Text = "00:00";
            this.lblRecordTimer.Visible = false;
            // 
            // lblStop
            // 
            this.lblStop.AutoSize = true;
            this.lblStop.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStop.ForeColor = System.Drawing.Color.RoyalBlue;
            this.lblStop.Location = new System.Drawing.Point(304, 83);
            this.lblStop.Name = "lblStop";
            this.lblStop.Size = new System.Drawing.Size(40, 16);
            this.lblStop.TabIndex = 2;
            this.lblStop.Text = "Stop";
            this.lblStop.Click += new System.EventHandler(this.lblStop_Click);
            // 
            // txtMessage
            // 
            this.txtMessage.Location = new System.Drawing.Point(6, 145);
            this.txtMessage.Name = "txtMessage";
            this.txtMessage.Size = new System.Drawing.Size(338, 43);
            this.txtMessage.TabIndex = 0;
            this.txtMessage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Peppermint_Outlook_AddIn.Properties.Resources.Recording_no_delay;
            this.pictureBox1.Location = new System.Drawing.Point(127, 13);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(100, 66);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 4;
            this.pictureBox1.TabStop = false;
            // 
            // PlayButton
            // 
            this.PlayButton.Image = global::Peppermint_Outlook_AddIn.Properties.Resources.play_2x;
            this.PlayButton.Location = new System.Drawing.Point(12, 99);
            this.PlayButton.Name = "PlayButton";
            this.PlayButton.Size = new System.Drawing.Size(25, 32);
            this.PlayButton.TabIndex = 9;
            this.PlayButton.TabStop = false;
            this.PlayButton.Visible = false;
            this.PlayButton.Click += new System.EventHandler(this.PlayButton_Click);
            // 
            // PauseButton
            // 
            this.PauseButton.Image = global::Peppermint_Outlook_AddIn.Properties.Resources.pause_2x;
            this.PauseButton.Location = new System.Drawing.Point(9, 99);
            this.PauseButton.Name = "PauseButton";
            this.PauseButton.Size = new System.Drawing.Size(25, 32);
            this.PauseButton.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.PauseButton.TabIndex = 10;
            this.PauseButton.TabStop = false;
            this.PauseButton.Visible = false;
            this.PauseButton.Click += new System.EventHandler(this.PauseButton_Click);
            // 
            // ProgressBar
            // 
            this.ProgressBar.BackColor = System.Drawing.Color.DarkRed;
            this.ProgressBar.ForeColor = System.Drawing.Color.Black;
            this.ProgressBar.Location = new System.Drawing.Point(44, 113);
            this.ProgressBar.Name = "ProgressBar";
            this.ProgressBar.Size = new System.Drawing.Size(296, 4);
            this.ProgressBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.ProgressBar.TabIndex = 11;
            this.ProgressBar.Value = 5;
            this.ProgressBar.Visible = false;
            // 
            // frmRecordAudio
            // 
            this.AcceptButton = this.btnAttachAudio;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(354, 248);
            this.Controls.Add(this.ProgressBar);
            this.Controls.Add(this.PauseButton);
            this.Controls.Add(this.lblStop);
            this.Controls.Add(this.lblRecordTimer);
            this.Controls.Add(this.txtMessage);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.btnAttachAudio);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.PlayButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmRecordAudio";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Peppermint Audio Recording";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmRecordAudio_FormClosing);
            this.Load += new System.EventHandler(this.frmRecordAudio_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.PlayButton)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.PauseButton)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnAttachAudio;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label txtMessage;
        private System.Windows.Forms.Label lblRecordTimer;
        private System.Windows.Forms.Label lblStop;
        private System.Windows.Forms.PictureBox PlayButton;
        private System.Windows.Forms.PictureBox PauseButton;
        private System.Windows.Forms.ProgressBar ProgressBar;
    }
}