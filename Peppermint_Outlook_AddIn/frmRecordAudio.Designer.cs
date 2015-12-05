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
            this.txtMessage = new System.Windows.Forms.Label();
            this.lblRecordTimer = new System.Windows.Forms.Label();
            this.lblStop = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.PlayButton = new System.Windows.Forms.PictureBox();
            this.txtTranscribedText = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.PlayButton)).BeginInit();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.ForeColor = System.Drawing.Color.White;
            this.btnCancel.Location = new System.Drawing.Point(3, 171);
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
            this.btnAttachAudio.Location = new System.Drawing.Point(179, 171);
            this.btnAttachAudio.Name = "btnAttachAudio";
            this.btnAttachAudio.Size = new System.Drawing.Size(169, 32);
            this.btnAttachAudio.TabIndex = 1;
            this.btnAttachAudio.Text = "Attach";
            this.btnAttachAudio.UseVisualStyleBackColor = false;
            this.btnAttachAudio.Click += new System.EventHandler(this.btnAttachAudio_Click);
            // 
            // txtMessage
            // 
            this.txtMessage.Location = new System.Drawing.Point(3, 114);
            this.txtMessage.Name = "txtMessage";
            this.txtMessage.Size = new System.Drawing.Size(294, 43);
            this.txtMessage.TabIndex = 5;
            this.txtMessage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblRecordTimer
            // 
            this.lblRecordTimer.AutoSize = true;
            this.lblRecordTimer.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRecordTimer.ForeColor = System.Drawing.Color.DarkGray;
            this.lblRecordTimer.Location = new System.Drawing.Point(162, 89);
            this.lblRecordTimer.Name = "lblRecordTimer";
            this.lblRecordTimer.Size = new System.Drawing.Size(36, 16);
            this.lblRecordTimer.TabIndex = 6;
            this.lblRecordTimer.Text = "0:00";
            this.lblRecordTimer.Visible = false;
            // 
            // lblStop
            // 
            this.lblStop.AutoSize = true;
            this.lblStop.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStop.ForeColor = System.Drawing.Color.RoyalBlue;
            this.lblStop.Location = new System.Drawing.Point(303, 127);
            this.lblStop.Name = "lblStop";
            this.lblStop.Size = new System.Drawing.Size(40, 16);
            this.lblStop.TabIndex = 7;
            this.lblStop.Text = "Stop";
            this.lblStop.Click += new System.EventHandler(this.lblStop_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Peppermint_Outlook_AddIn.Properties.Resources.Recording_no_delay;
            this.pictureBox1.Location = new System.Drawing.Point(131, 15);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(100, 66);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 4;
            this.pictureBox1.TabStop = false;
            // 
            // PlayButton
            // 
            this.PlayButton.Image = global::Peppermint_Outlook_AddIn.Properties.Resources.PlayButton;
            this.PlayButton.Location = new System.Drawing.Point(306, 114);
            this.PlayButton.Name = "PlayButton";
            this.PlayButton.Size = new System.Drawing.Size(39, 43);
            this.PlayButton.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.PlayButton.TabIndex = 9;
            this.PlayButton.TabStop = false;
            this.PlayButton.Visible = false;
            this.PlayButton.Click += new System.EventHandler(this.PlayButton_Click);
            // 
            // txtTranscribedText
            // 
            this.txtTranscribedText.Location = new System.Drawing.Point(6, 209);
            this.txtTranscribedText.Multiline = true;
            this.txtTranscribedText.Name = "txtTranscribedText";
            this.txtTranscribedText.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtTranscribedText.Size = new System.Drawing.Size(342, 67);
            this.txtTranscribedText.TabIndex = 10;
            // 
            // frmRecordAudio
            // 
            this.AcceptButton = this.btnAttachAudio;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(354, 288);
            this.Controls.Add(this.txtTranscribedText);
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
        private System.Windows.Forms.TextBox txtTranscribedText;
    }
}