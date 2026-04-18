using NAudio.CoreAudioApi;
using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;

namespace XJAudioLatencyTester
{
    /// <summary>
    /// Measures audio round-trip latency using cross-correlation of a chirp signal.
    ///
    /// Algorithm:
    ///   1. Optionally plays a 1s low-amplitude warm-up tone (for Bluetooth wake-up).
    ///   2. Plays a known chirp (200 Hz→8 kHz, 100 ms) at full amplitude.
    ///   3. Records simultaneously, noting the exact Stopwatch tick when recording and
    ///      playback each start — this compensates for the recording→playback time offset.
    ///   4. Cross-correlates the reference chirp against the recording over a ±50/+800 ms
    ///      search window using a coarse (0.25 ms step) + fine (per-sample) pass.
    ///   5. The peak position gives the precise delay with ~0.02 ms (1 sample) resolution.
    /// </summary>
    public class ImprovedLatencyMeasurer
    {
        private const int SampleRate = 44100;
        private const double WarmupDurationSec = 1.0;
        private const double ChirpDurationSec = 0.1;    // 100 ms → 4 410 samples
        private const double ChirpFreqStart = 200.0;
        private const double ChirpFreqEnd = 8000.0;
        private const double WarmupAmplitude = 0.05;
        private const double ChirpAmplitude = 0.95;
        private const double MinCorrelation = 0.20;     // reject peak below this

        private readonly float threshold;
        private readonly Action<string> log;
        private readonly bool bluetoothWarmup;
        private readonly int testDurationMs;
        private readonly object dataLock = new object();

        private List<byte> recordedData;
        private float[] referenceChirp;   // normalized [-1..1] float samples

        public ImprovedLatencyMeasurer(float detectionThreshold, Action<string> logger,
            bool enableBluetoothWarmup = true, int testDurationMs = 5000)
        {
            threshold = detectionThreshold;
            log = logger;
            bluetoothWarmup = enableBluetoothWarmup;
            this.testDurationMs = testDurationMs;
        }

        public (double latency, byte[] recordedData) MeasureLatency(string wasapiDeviceId, int wmmDeviceId, int inputDeviceId)
        {
            this.recordedData = new List<byte>(SampleRate * 2 * 4);

            // Look up MMDevice on THIS thread (worker) — passing MMDevice across STA/MTA causes E_NOINTERFACE.
            var enumerator = new MMDeviceEnumerator();
            var outputDevice = enumerator.GetDevice(wasapiDeviceId);

            log($"Starting measurement - Output: {outputDevice.FriendlyName}, Input: {inputDeviceId}");
            log($"Detection threshold: {threshold:F6} (amplitude) = {20 * Math.Log10(threshold):F1} dB");

            double effectiveWarmupSec = bluetoothWarmup ? WarmupDurationSec : 0.0;
            if (bluetoothWarmup)
                log($"Bluetooth warm-up: {WarmupDurationSec}s init + {ChirpDurationSec}s chirp (200 Hz→8 kHz)");
            else
                log($"No warm-up: chirp starts immediately");

            var waveFormat = new WaveFormat(SampleRate, 16, 1);

            var waveIn = new WaveInEvent
            {
                DeviceNumber = inputDeviceId,
                WaveFormat = waveFormat,
                BufferMilliseconds = 10
            };
            waveIn.DataAvailable += OnDataAvailable;

            var testSignal = GenerateTestSignal(effectiveWarmupSec);
            log($"Generated test signal: {testSignal.Length / 2} samples ({testSignal.Length} bytes)");

            // Source provider — always 44100/16/1; may be resampled for WASAPI below.
            var rawProvider = new BufferedWaveProvider(waveFormat) { BufferDuration = TimeSpan.FromSeconds(5) };

            // Build the output player: WASAPI shared (format-aware) → WinMM fallback.
            IWavePlayer waveOut = CreateOutputPlayer(outputDevice, wmmDeviceId, rawProvider, waveFormat);

            var sw = Stopwatch.StartNew();
            long recStartTick, playStartTick;

            try
            {
                waveIn.StartRecording();
                recStartTick = sw.ElapsedTicks;
                log("Recording started");

                rawProvider.AddSamples(testSignal, 0, testSignal.Length);
                waveOut.Play();
                playStartTick = sw.ElapsedTicks;
                log("Playback started");

                // How many ms elapsed between recording start and playback start.
                // The chirp appears in the recording at (effectiveWarmupSec*1000 + recToPlayMs) ms
                // from byte 0 of the recording — even if latency were 0.
                double recToPlayMs = (playStartTick - recStartTick) * 1000.0 / Stopwatch.Frequency;
                log($"Recording→Playback offset: {recToPlayMs:F2}ms");

                // Wait until we have enough data:
                // warmup + chirp + 800ms search headroom, plus the recording offset
                int minBytes = BytesFromMs(effectiveWarmupSec * 1000 + ChirpDurationSec * 1000 + 800 + recToPlayMs);
                int waited = 0;
                while (waited < testDurationMs)
                {
                    Thread.Sleep(50);
                    waited += 50;
                    int count;
                    lock (dataLock) { count = this.recordedData.Count; }
                    if (count >= minBytes) break;
                }

                byte[] snapshot;
                lock (dataLock) { snapshot = this.recordedData.ToArray(); }
                log($"Captured {snapshot.Length} bytes ({snapshot.Length / 2.0 / SampleRate * 1000:F0}ms), analysing...");

                double latencyMs = FindLatencyByCrossCorrelation(snapshot, effectiveWarmupSec, recToPlayMs);
                return (latencyMs, snapshot);
            }
            finally
            {
                try
                {
                    waveIn.StopRecording();
                    waveOut.Stop();
                    waveIn.Dispose();
                    waveOut.Dispose();
                    log("Audio devices disposed");
                }
                catch { }
            }
        }

        // ── Output player factory ─────────────────────────────────────────────────

        // Tries WASAPI shared mode with automatic format conversion (handles 48kHz/float devices).
        // Falls back to WinMM if WASAPI is unavailable (e.g., device locked by Voicemeeter).
        private IWavePlayer CreateOutputPlayer(MMDevice wasapiDevice, int wmmDeviceId,
            BufferedWaveProvider rawProvider, WaveFormat sourceFormat)
        {
            // ── WASAPI attempt ───────────────────────────────────────────────────
            try
            {
                WaveFormat mixFormat;
                try { mixFormat = wasapiDevice.AudioClient.MixFormat; }
                catch { mixFormat = sourceFormat; }

                IWaveProvider wasapiSource = rawProvider;
                if (mixFormat.SampleRate != sourceFormat.SampleRate
                    || mixFormat.Channels != sourceFormat.Channels
                    || mixFormat.BitsPerSample != sourceFormat.BitsPerSample)
                {
                    log($"Output format: {sourceFormat.SampleRate}Hz/16bit/1ch " +
                        $"→ {mixFormat.SampleRate}Hz/{mixFormat.BitsPerSample}bit/{mixFormat.Channels}ch");
                    wasapiSource = new MediaFoundationResampler(rawProvider, mixFormat) { ResamplerQuality = 60 };
                }

                var wasapiOut = new WasapiOut(wasapiDevice, AudioClientShareMode.Shared, true, 200);
                wasapiOut.Init(wasapiSource);
                log("Output: WASAPI shared mode");
                return wasapiOut;
            }
            catch (Exception ex)
            {
                log($"WASAPI unavailable ({ex.Message.Split('\n')[0].Trim()}), trying WinMM...");
            }

            // ── WinMM fallback ───────────────────────────────────────────────────
            try
            {
                var mmOut = new WaveOutEvent { DeviceNumber = wmmDeviceId };
                mmOut.Init(rawProvider);
                log("Output: WinMM (WaveOutEvent)");
                return mmOut;
            }
            catch (Exception ex)
            {
                throw new Exception(
                    $"Device unavailable via both WASAPI and WinMM. " +
                    $"If Voicemeeter is running, physical outputs are locked — use 'VoiceMeeter Input' instead. " +
                    $"({ex.Message.Split('\n')[0].Trim()})");
            }
        }

        // ── Signal generation ────────────────────────────────────────────────────

        private byte[] GenerateTestSignal(double warmupSec)
        {
            double totalSec = warmupSec + ChirpDurationSec + 0.9; // tail silence
            int totalSamples = (int)(SampleRate * totalSec);
            var buffer = new byte[totalSamples * 2];

            int warmupSamples = (int)(SampleRate * warmupSec);
            int chirpSamples = (int)(SampleRate * ChirpDurationSec);

            // Warm-up: low-amplitude 440 Hz sine
            if (warmupSamples > 0)
            {
                double amp = WarmupAmplitude * short.MaxValue;
                for (int i = 0; i < warmupSamples; i++)
                {
                    double t = (double)i / SampleRate;
                    Write16(buffer, i, (short)(amp * Math.Sin(2 * Math.PI * 440 * t)));
                }
            }

            // Chirp: linear frequency sweep, captured as reference
            double k = (ChirpFreqEnd - ChirpFreqStart) / ChirpDurationSec;
            double chirpAmp = ChirpAmplitude * short.MaxValue;
            referenceChirp = new float[chirpSamples];

            for (int i = 0; i < chirpSamples; i++)
            {
                double t = (double)i / SampleRate;
                double phase = 2 * Math.PI * (ChirpFreqStart * t + 0.5 * k * t * t);
                float fSample = (float)Math.Sin(phase);
                referenceChirp[i] = fSample;
                Write16(buffer, warmupSamples + i, (short)(chirpAmp * fSample));
            }

            return buffer;
        }

        // ── Cross-correlation analysis ────────────────────────────────────────────

        private double FindLatencyByCrossCorrelation(byte[] recording, double warmupSec, double recToPlayMs)
        {
            // Expected position of the chirp in the recording (assuming 0 latency device)
            double expectedChirpMs = warmupSec * 1000.0 + recToPlayMs;
            int expectedByte = BytesFromMs(expectedChirpMs);

            log($"Expected chirp start in recording: {expectedChirpMs:F1}ms (byte {expectedByte})");

            // Search window: 50ms before expected (handles tiny timing overestimate)
            //                up to 800ms after expected (handles very high latency)
            int searchFrom = Math.Max(0, expectedByte - BytesFromMs(50));
            int searchTo = Math.Min(
                recording.Length - referenceChirp.Length * 2,
                expectedByte + BytesFromMs(800));

            if (searchTo <= searchFrom)
                throw new Exception(
                    $"Not enough data for cross-correlation. " +
                    $"Have {recording.Length} bytes, need ~{expectedByte + referenceChirp.Length * 2} bytes.");

            // Precompute ‖reference‖ (constant for all positions)
            double refNorm = Math.Sqrt(referenceChirp.Sum(s => (double)s * s));

            log($"Cross-correlating over [{searchFrom / 2.0 / SampleRate * 1000:F0}ms – {searchTo / 2.0 / SampleRate * 1000:F0}ms] in recording");

            // Coarse pass: step = 11 samples ≈ 0.25 ms
            const int CoarseStepSamples = 11;
            int coarseStep = CoarseStepSamples * 2;
            double bestCorr = double.MinValue;
            int bestByte = searchFrom;

            for (int pos = searchFrom; pos <= searchTo; pos += coarseStep)
            {
                double c = NormalizedCorrelationAt(recording, pos, refNorm);
                if (c > bestCorr) { bestCorr = c; bestByte = pos; }
            }

            log($"Coarse peak: corr={bestCorr:F3} at {bestByte / 2.0 / SampleRate * 1000:F1}ms");

            // Fine pass: ±3 ms around coarse peak, per-sample step
            int fineFrom = Math.Max(searchFrom, bestByte - BytesFromMs(3));
            int fineTo = Math.Min(searchTo, bestByte + BytesFromMs(3));
            bestCorr = double.MinValue;

            for (int pos = fineFrom; pos <= fineTo; pos += 2)
            {
                double c = NormalizedCorrelationAt(recording, pos, refNorm);
                if (c > bestCorr) { bestCorr = c; bestByte = pos; }
            }

            log($"Fine peak: corr={bestCorr:F3} at byte {bestByte} ({bestByte / 2.0 / SampleRate * 1000:F2}ms)");

            if (bestCorr < MinCorrelation)
                throw new Exception(
                    $"Cross-correlation peak too weak ({bestCorr:F3}). " +
                    $"Check microphone placement and volume.");

            double detectedMs = bestByte / 2.0 / SampleRate * 1000.0;
            double latencyMs = detectedMs - expectedChirpMs;

            log($"Chirp detected at {detectedMs:F2}ms in recording");
            log($"Expected at {expectedChirpMs:F1}ms (zero-latency baseline)");
            log($"Measured latency: {latencyMs:F1}ms");

            return latencyMs;
        }

        // Normalized cross-correlation coefficient at one position (scalar in [-1, 1])
        private double NormalizedCorrelationAt(byte[] rec, int startByte, double refNorm)
        {
            int len = referenceChirp.Length;
            if (startByte + len * 2 > rec.Length) return 0.0;

            double sumXY = 0, sumX2 = 0;
            for (int i = 0; i < len; i++)
            {
                float recSample = (short)((rec[startByte + i * 2 + 1] << 8) | rec[startByte + i * 2]) / 32768f;
                float refSample = referenceChirp[i];
                sumXY += recSample * refSample;
                sumX2 += recSample * recSample;
            }

            if (sumX2 < 1e-10) return 0.0;
            return sumXY / (Math.Sqrt(sumX2) * refNorm);
        }

        // ── Helpers ───────────────────────────────────────────────────────────────

        private void OnDataAvailable(object sender, WaveInEventArgs e)
        {
            byte[] buf = new byte[e.BytesRecorded];
            Array.Copy(e.Buffer, buf, e.BytesRecorded);
            lock (dataLock) { recordedData.AddRange(buf); }
        }

        private static int BytesFromMs(double ms) => (int)(ms / 1000.0 * SampleRate) * 2;

        private static void Write16(byte[] buf, int sampleIndex, short value)
        {
            buf[sampleIndex * 2] = (byte)(value & 0xFF);
            buf[sampleIndex * 2 + 1] = (byte)(value >> 8);
        }
    }

    // Helper classes kept for any external references
    public class ChunkAnalysisResult
    {
        public float MaxAmplitude { get; set; }
        public double MaxDB { get; set; }
        public int SamplesAboveThreshold { get; set; }
        public int TotalSamples { get; set; }
    }

    public class DataAnalysis
    {
        public List<ChunkAnalysis> Chunks { get; set; } = new List<ChunkAnalysis>();
    }

    public class ChunkAnalysis
    {
        public double Time { get; set; }
        public double MaxDB { get; set; }
        public int SamplesAboveThreshold { get; set; }
    }
}
