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

            outputFilename = String.Format("Peppermint_Message {0:yyyy-MMM-dd HH-mm-ss}.wav", DateTime.Now);
            writer = new WaveFileWriter(Path.Combine(outputFolder, outputFilename), waveIn.WaveFormat);
            try 
            { 
                waveIn.StartRecording();
            }
            catch (Exception ex)
            {
                string msg = String.Format("Your microphone is disabled, please check your settings.");
                MessageBox.Show(msg,"Error while recording audio", MessageBoxButtons.OK,MessageBoxIcon.Error);
                this.Dispose();
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
                    MessageBox.Show("There is a problem with your microphone. Please check it and try again", "A problem was encountered during recording", MessageBoxButtons.OK,MessageBoxIcon.Error);
                    ThisAddIn.AttachmentFilePath = String.Empty;
                    this.Dispose();
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
