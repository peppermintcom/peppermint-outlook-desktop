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
        private string MIC_ERRROR = "Your microphone is not working. Please check your audio settings and try again.";

        public frmRecordAudio()
        {
            InitializeComponent();

            outputFolder = Path.Combine(Path.GetTempPath(), "Peppermint_Outlook_Addin");
            Directory.CreateDirectory(outputFolder);
        }

        private void frmRecordAudio_Load(object sender, EventArgs e)
        {
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
                pictureBox1.Image = Properties.Resources.GIF_01;
            }
            catch (Exception ex)
            {
                ThisAddIn.AttachmentFilePath = String.Empty;
                txtMessage.Text = MIC_ERRROR;
                pictureBox1.Image = Properties.Resources.icon_mic_off;
                waveIn = null;
            }
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
                    txtMessage.Text = MIC_ERRROR;
                    pictureBox1.Image = Properties.Resources.icon_mic_off;
                }
            }
        }

        private void FinalizeWaveFile()
        {
            if (writer != null)
            {
                writer.Dispose();
                writer = null;
            }
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
            if (waveIn != null)
            {
                waveIn.StopRecording();
                ThisAddIn.AttachmentFilePath = outputFolder + "\\" + outputFilename;

                FinalizeWaveFile();
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
    }
}
