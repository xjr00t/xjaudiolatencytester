using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Net.Configuration;
using System.Threading;

namespace XJAudioLatencyTester
{
    public class ThresholdCalibrator
    {
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

            waveIn.RecordingStopped += (s, e) =>
            {
                recordingComplete = true;
            };

            try
            {
                // Record background noise for 2 seconds
                waveIn.StartRecording();
                Thread.Sleep(2000);
                waveIn.StopRecording();

                // Wait for recording to stop
                int timeout = 1000;
                while (!recordingComplete && timeout > 0)
                {
                    Thread.Sleep(100);
                    timeout -= 100;
                }

                // Calculate RMS of background noise
                double sum = 0;
                int sampleCount = 0;
                float avg = 0;
                float max = 0;

                for (int i = 0; i < recordedData.Count - 1; i += 2)
                {
                    short sample = (short)((recordedData[i + 1] << 8) | recordedData[i + 0]);
                    var sample32 = sample / 32768f;
                    if (sample32 < 0) sample32 = -sample32;
                    //sum += sample32;
                    if (sample32 > max) max = sample32;
                    sampleCount++;
                }

                //if (sampleCount > 0) avg = (float)(sum / sampleCount);


                //double rms = Math.Sqrt(sumSquares / sampleCount);

                // Set threshold to 3x RMS (common practice)
                //float threshold = (float)(rms * 3);

                // Clamp to reasonable values
                //threshold = Math.Max(0.001f, Math.Min(0.5f, threshold));
                //int threshold = (int)(avg * 100);
                float thresholdFloat = 10000f * max;
                int threshold = (int)thresholdFloat;
                return threshold;
            }
            finally
            {
                waveIn.Dispose();
            }
        }
    }
}