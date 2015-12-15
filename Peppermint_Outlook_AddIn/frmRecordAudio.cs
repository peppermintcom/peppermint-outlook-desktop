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
using System.Threading;

using System.Speech.Recognition;

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
        private const int MAX_RECORDING_TIME = 10 * 60; // max audio recording time seconds = 10 mins
        private string RECORDING_CONCLUDED = "Recording concluded";
        private string PLAYING_AUDIO = "Playing recorded message ...";

        private bool bRecordingInProgress;

        private SpeechRecognitionEngine _recognizer;

        DirectSoundOut audioOutput;
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
                            lblStop.Visible = false;
                        }
                        bRecordingInProgress = false;
                        lblRecordTimer.Visible = false;
                    }
                    catch (Exception)
                    {
                        bRecordingInProgress = false;
                        lblRecordTimer.Visible = false;
                        lblStop.Visible = false;
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
            lblRecordTimer.Visible = true;

            outputFolder = Path.Combine(Path.GetTempPath(), "Peppermint_Outlook_Addin");
            Directory.CreateDirectory(outputFolder);
        }

        private void StartRecording()
        {
            bRecordingInProgress = true;
            lblRecordTimer.Visible = true;
            ThisAddIn.PEPPERMINT_TRANSCRIBED_AUDIO = String.Empty;

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
            catch (Exception)
            {
                bRecordingInProgress = false;
                lblRecordTimer.Visible = false;
                lblStop.Visible = false;
                ThisAddIn.AttachmentFilePath = String.Empty;
                txtMessage.Text = MIC_ERROR;
                pictureBox1.Image = Properties.Resources.icon_mic_off;
                waveIn = null;
                btnCancel.Enabled = false;
                btnAttachAudio.Enabled = false;
            }

            _recognizer = new SpeechRecognitionEngine(new System.Globalization.CultureInfo("en-US"));

            try
            {
                _recognizer.LoadGrammar(new DictationGrammar());

                _recognizer.SpeechRecognized +=_recognizer_SpeechRecognized;

                _recognizer.SetInputToDefaultAudioDevice(); // set the input of the speech recognizer to the default audio device
                _recognizer.RecognizeAsync(RecognizeMode.Multiple); // recognize speech asynchronous
            }

            catch (InvalidOperationException exception)
            {
                string msg = String.Format("Could not recognize input from default audio device. Is a microphone or sound card available?\r\n{0} - {1}.", exception.Source, exception.Message);
                MessageBox.Show(msg);
            }
        }

        void _recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            //txtTranscribedText.AppendText(e.Result.Text + " ");
            //ThisAddIn.PEPPERMINT_TRANSCRIBED_AUDIO = txtTranscribedText.Text;

            ThisAddIn.PEPPERMINT_TRANSCRIBED_AUDIO += e.Result.Text + " ";
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
                    lblStop.Visible = false;
                    pictureBox1.Image = Properties.Resources.icon_mic_off;
                    waveIn = null;
                    btnCancel.Enabled = false;
                    btnAttachAudio.Enabled = false;
                }
            }
            bRecordingInProgress = false;
            lblRecordTimer.Visible = false;
        }

        private void FinalizeWaveFile()
        {
            if (writer != null)
            {
                writer.Dispose();
                writer = null;
            }
            bRecordingInProgress = false;
            lblRecordTimer.Visible = false;
        }

        void waveIn_DataAvailable(object sender, WaveInEventArgs e)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new EventHandler<WaveInEventArgs>(waveIn_DataAvailable), sender, e);
            }
            else
            {
                if (writer == null)
                    return;

                writer.Write(e.Buffer, 0, e.BytesRecorded);
                int secondsRecorded = (int)(writer.Length / writer.WaveFormat.AverageBytesPerSecond);

                TimeSpan ts = TimeSpan.FromSeconds(secondsRecorded);
                lblRecordTimer.Text = string.Format("{0:D2}:{1:D2}", ts.Minutes,ts.Seconds);

                if (secondsRecorded >= MAX_RECORDING_TIME)
                {
                    btnAttachAudio_Click(sender, e);
                    this.DialogResult = DialogResult.OK;
                    this.Dispose();
                }
            }
        }

        private void btnAttachAudio_Click(object sender, EventArgs e)
        {
            // Either the button has text "Attach" or "Record". If it is "Attach" just complete the recording and attach the file,
            // ELSE, if it is "Record" start a new recording session
            if (btnAttachAudio.Text == "Attach")
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
                btnAttachAudio.Text = "Attach";
                this.DialogResult = DialogResult.None;
                StartRecording();
                lblStop.Visible = true;
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

        private void lblStop_Click(object sender, EventArgs e)
        {
            lblStop.Visible = false;
            PlayButton.Visible = true;
            //ProgressBar.Visible = true;
            PauseButton.Visible = true;

            // Stop the recording, but do not attach the file, yet
            if (waveIn != null)
            {
                waveIn.StopRecording();

                FinalizeWaveFile();

                pictureBox1.Image = Properties.Resources.Logo;
                txtMessage.Text = RECORDING_CONCLUDED;
            }
            if ( _recognizer != null )
                _recognizer.SpeechRecognized -= _recognizer_SpeechRecognized;
        }

        private void PlayButton_Click(object sender, EventArgs e)
        {
            txtMessage.Text = PLAYING_AUDIO;
            lblStop.Visible = false;
            //this.Enabled = false;

            string strFileToPlay = outputFolder + "\\" + outputFilename;

            
            if (audioOutput != null ) // Playback is already in progress i.e. re-play after playback has been paused
            {
                audioOutput.Play();
                return;
            }
            
            // If the recorded file is present then play the attachment
            if (File.Exists(strFileToPlay))
            {
                var soundFile = strFileToPlay;
                var wfr = new WaveFileReader(soundFile);
                WaveChannel32 wc = new WaveChannel32(wfr) { PadWithZeroes = false };
                audioOutput = new DirectSoundOut();
                {
                    audioOutput.PlaybackStopped += audioOutput_PlaybackStopped;
                    audioOutput.Init(wc);

                    audioOutput.Play();
                }
            }
            else
            {
                MessageBox.Show("Could not find the recorded file\n\nPlease try recording again", "Audio file not found", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            this.Enabled = true;
        }

        void audioOutput_PlaybackStopped(object sender, StoppedEventArgs e)
        {
            audioOutput.Dispose();
            audioOutput = null;
        }

        private void PauseButton_Click(object sender, EventArgs e)
        {
            if (audioOutput != null)
                audioOutput.Pause();
        }
    }
}
