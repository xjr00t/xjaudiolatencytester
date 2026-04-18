using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Threading;

namespace XJAudioLatencyTester
{
    public class ThresholdCalibrator
    {
        /// <summary>
        /// Records 2 seconds of background noise and returns a detection threshold
        /// set ~14 dB above the measured noise floor (5x peak amplitude).
        /// Returns an amplitude value in [0.001, 0.5].
        /// </summary>
        public float Calibrate(int inputDeviceId)
        {
            var recordedData = new List<byte>();
            var recordingComplete = false;

            var waveIn = new WaveInEvent
            {
                DeviceNumber = inputDeviceId,
                WaveFormat = new WaveFormat(44100, 16, 1),
                BufferMilliseconds = 100
            };

            waveIn.DataAvailable += (s, e) =>
            {
                byte[] buffer = new byte[e.BytesRecorded];
                Array.Copy(e.Buffer, buffer, e.BytesRecorded);
                recordedData.AddRange(buffer);
            };

            waveIn.RecordingStopped += (s, e) => { recordingComplete = true; };

            try
            {
                waveIn.StartRecording();
                Thread.Sleep(2000);
                waveIn.StopRecording();

                int timeout = 1000;
                while (!recordingComplete && timeout > 0) { Thread.Sleep(50); timeout -= 50; }

                // Find peak amplitude of background noise
                float peak = 0;
                for (int i = 0; i < recordedData.Count - 1; i += 2)
                {
                    short sample = (short)((recordedData[i + 1] << 8) | recordedData[i]);
                    float amplitude = Math.Abs(sample / 32768f);
                    if (amplitude > peak) peak = amplitude;
                }

                // Threshold = noise peak × 5  (~14 dB headroom above noise floor).
                // Clamp to [0.001, 0.5] so it stays well below the chirp signal (~0.95).
                float threshold = Math.Max(0.001f, Math.Min(0.5f, peak * 5f));
                return threshold;
            }
            finally
            {
                waveIn.Dispose();
            }
        }
    }
}
