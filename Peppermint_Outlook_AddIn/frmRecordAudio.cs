using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;

using NAudio.Wave;

namespace Peppermint_Outlook_AddIn
{
    public partial class frmRecordAudio : Form
    {
        private IWaveIn waveIn;
        private WaveFileWriter writer;
        private string outputFilename;
        private readonly string outputFolder;
        private string RECORDING = "Recording your message ...";
        private string MIC_ERROR = "Your microphone is not working. Please check your audio settings and try again.";
        private string MIC_INSERTED = "Ok The problem seems to be fixed, click on Record when ready";

        private bool bRecordingInProgress;

        // WIN API
        private const int WM_DEVICECHANGE = 0x0219;
        
        // New device has been plugged in
        private const int DBT_DEVICEARRIVAL = 0x8000;
        
        // Device removed 
        private const int DBT_DEVICEREMOVECOMPLETE = 0x8004;
        
        // Device has changed
        private const int DBT_DEVNODES_CHANGED = 0x0007;
        
        protected override void WndProc(ref Message m)
        {
            if (bRecordingInProgress == true)
            {
                // Don't process any device change events now. mic removal will anyway will be catched as an exception while recording.
            }
            else 
            {
                if (m.Msg == WM_DEVICECHANGE)
                {
                    if (waveIn == null)
                    {
                        waveIn = new WaveIn();
                        waveIn.WaveFormat = new WaveFormat(8000, 1);
                    }

                    try
                    {
                        waveIn.StartRecording();
                        waveIn.StopRecording();
                        waveIn.Dispose();
                        waveIn = null;

                        if(btnAttachAudio.Enabled == false)
                        {
                            btnAttachAudio.Enabled = true;
                            btnAttachAudio.Text = "Record";
                            btnCancel.Enabled = true;
                            pictureBox1.Image = Properties.Resources.icon_mic_on;
                            txtMessage.Text = MIC_INSERTED;
                        }
                        bRecordingInProgress = false;
                    }
                    catch (Exception ex)
                    {
                        bRecordingInProgress = false;
                        ThisAddIn.AttachmentFilePath = String.Empty;
                        txtMessage.Text = MIC_ERROR;
                        pictureBox1.Image = Properties.Resources.icon_mic_off;
                        waveIn = null;
                        btnCancel.Enabled = false;
                        btnAttachAudio.Enabled = false;
                    }
                }
            }

            base.WndProc(ref m);
        }
        public frmRecordAudio()
        {
            InitializeComponent();

            bRecordingInProgress = true;

            outputFolder = Path.Combine(Path.GetTempPath(), "Peppermint_Outlook_Addin");
            Directory.CreateDirectory(outputFolder);
        }

        private void StartRecording()
        {
            bRecordingInProgress = true;

            if (waveIn == null)
            {
                waveIn = new WaveIn();
                waveIn.WaveFormat = new WaveFormat(8000, 1);

                waveIn.DataAvailable += waveIn_DataAvailable;
                waveIn.RecordingStopped += waveIn_RecordingStopped;
            }

            outputFilename = String.Format("Peppermint_Message {0:yyyy-MMM-dd h-mm-ss tt}.wav", DateTime.Now);
            writer = new WaveFileWriter(Path.Combine(outputFolder, outputFilename), waveIn.WaveFormat);
            try
            {
                waveIn.StartRecording();
                txtMessage.Text = RECORDING;
                pictureBox1.Image = Properties.Resources.Recording_no_delay;
            }
            catch (Exception ex)
            {
                bRecordingInProgress = false;
                ThisAddIn.AttachmentFilePath = String.Empty;
                txtMessage.Text = MIC_ERROR;
                pictureBox1.Image = Properties.Resources.icon_mic_off;
                waveIn = null;
                btnCancel.Enabled = false;
                btnAttachAudio.Enabled = false;
            }

        }

        private void frmRecordAudio_Load(object sender, EventArgs e)
        {
            StartRecording();
        }

        void waveIn_RecordingStopped(object sender, StoppedEventArgs e)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new EventHandler<StoppedEventArgs>(waveIn_RecordingStopped), sender, e);
            }
            else
            {
                FinalizeWaveFile();
                if (e.Exception != null)
                {
                    ThisAddIn.AttachmentFilePath = String.Empty;
                    txtMessage.Text = MIC_ERROR;
                    pictureBox1.Image = Properties.Resources.icon_mic_off;
                    waveIn = null;
                    btnCancel.Enabled = false;
                    btnAttachAudio.Enabled = false;
                }
            }
            bRecordingInProgress = false;
        }

        private void FinalizeWaveFile()
        {
            if (writer != null)
            {
                writer.Dispose();
                writer = null;
            }
            bRecordingInProgress = false;
        }

        void waveIn_DataAvailable(object sender, WaveInEventArgs e)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new EventHandler<WaveInEventArgs>(waveIn_DataAvailable), sender, e);
            }
            else
            {
                writer.Write(e.Buffer, 0, e.BytesRecorded);
                int secondsRecorded = (int)(writer.Length / writer.WaveFormat.AverageBytesPerSecond);
            }
        }

        private void btnAttachAudio_Click(object sender, EventArgs e)
        {
            // Either the button has text "Done" or "Record". If it is "Done" just complete the recording and attach the file,
            // ELSE, if it is "Record" start a new recording session
            if (btnAttachAudio.Text == "Done")
            { 
                if (waveIn != null)
                {
                    waveIn.StopRecording();
                    ThisAddIn.AttachmentFilePath = outputFolder + "\\" + outputFilename;

                    FinalizeWaveFile();
                }
            }
            if (btnAttachAudio.Text == "Record")
            {
                // Start the recording
                btnAttachAudio.Text = "Done";
                this.DialogResult = DialogResult.None;
                StartRecording();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (waveIn != null)
            {
                waveIn.StopRecording();
                ThisAddIn.AttachmentFilePath = String.Empty;
            }
        }

        private void frmRecordAudio_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.DialogResult == DialogResult.None)
                e.Cancel = true;
        }
    }
}
