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
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.txtMessage = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.ForeColor = System.Drawing.Color.White;
            this.btnCancel.Location = new System.Drawing.Point(10, 151);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(147, 32);
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
            this.btnAttachAudio.Location = new System.Drawing.Point(164, 151);
            this.btnAttachAudio.Name = "btnAttachAudio";
            this.btnAttachAudio.Size = new System.Drawing.Size(147, 32);
            this.btnAttachAudio.TabIndex = 1;
            this.btnAttachAudio.Text = "Done";
            this.btnAttachAudio.UseVisualStyleBackColor = false;
            this.btnAttachAudio.Click += new System.EventHandler(this.btnAttachAudio_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Peppermint_Outlook_AddIn.Properties.Resources.icon_mic_on;
            this.pictureBox1.Location = new System.Drawing.Point(112, 15);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(100, 66);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 4;
            this.pictureBox1.TabStop = false;
            // 
            // txtMessage
            // 
            this.txtMessage.Location = new System.Drawing.Point(24, 94);
            this.txtMessage.Name = "txtMessage";
            this.txtMessage.Size = new System.Drawing.Size(273, 43);
            this.txtMessage.TabIndex = 5;
            this.txtMessage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // frmRecordAudio
            // 
            this.AcceptButton = this.btnAttachAudio;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(322, 190);
            this.Controls.Add(this.txtMessage);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.btnAttachAudio);
            this.Controls.Add(this.btnCancel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmRecordAudio";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Peppermint record audio";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmRecordAudio_FormClosing);
            this.Load += new System.EventHandler(this.frmRecordAudio_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnAttachAudio;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label txtMessage;
    }
}