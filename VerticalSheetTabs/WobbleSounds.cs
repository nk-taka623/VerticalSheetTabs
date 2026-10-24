using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;

namespace VerticalSheetTabs
{
    // The sound effects of the wobble animations (Agents.md 12.172, 12.174): one short sound for each motion, embedded in the
    // add-in as 16 bit mono wave files (Sounds\*.wav). Every sound plays on a voice of its own (the Windows waveOut API), so
    // a sound that starts while another one is still playing is mixed over it by Windows. The samples get the volume, a
    // short fade in and a long fade out before they are played.
    public static class WobbleSounds
    {
        public const int VolumeDefault = 60;
        public const int FadeOutMs = 200; // the end of every sound dies away over this time
        private const int FadeInMs = 5;
        private const int TailMs = 40;    // silence after the sound, so the device does not cut the fade
        private const int MaxVoices = 8;
        private const int SampleRate = 44100;

        private static readonly Dictionary<WobbleType, string> Resources = new Dictionary<WobbleType, string>
        {
            { WobbleType.Shake, "shake" },
            { WobbleType.Wobble, "wobble" },
            { WobbleType.Jelly, "jelly" },
            { WobbleType.Bounce, "bounce" },
            { WobbleType.Pulse, "pulse" },
            { WobbleType.Swing, "swing" }
        };

        private static readonly Dictionary<WobbleType, short[]> Cache = new Dictionary<WobbleType, short[]>();
        private static readonly List<Voice> Voices = new List<Voice>();
        private static readonly object Lock = new object();
        private static readonly System.Collections.Concurrent.ConcurrentQueue<long> DoneIds = new System.Collections.Concurrent.ConcurrentQueue<long>();
        private static long _nextId;

        // ---- the waveOut API ----

        private const int WaveMapper = -1;
        private const int CallbackFunction = 0x00030000;
        private const int WomDone = 0x3BD;

        [StructLayout(LayoutKind.Sequential)]
        private struct WaveFormat
        {
            public ushort wFormatTag;
            public ushort nChannels;
            public uint nSamplesPerSec;
            public uint nAvgBytesPerSec;
            public ushort nBlockAlign;
            public ushort wBitsPerSample;
            public ushort cbSize;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct WaveHeader
        {
            public IntPtr lpData;
            public uint dwBufferLength;
            public uint dwBytesRecorded;
            public IntPtr dwUser;
            public uint dwFlags;
            public uint dwLoops;
            public IntPtr lpNext;
            public IntPtr reserved;
        }

        private delegate void WaveCallback(IntPtr hwo, uint message, IntPtr instance, IntPtr param1, IntPtr param2);

        [DllImport("winmm.dll")]
        private static extern int waveOutOpen(out IntPtr hwo, int deviceId, ref WaveFormat format, WaveCallback callback, IntPtr instance, int flags);

        [DllImport("winmm.dll")]
        private static extern int waveOutPrepareHeader(IntPtr hwo, IntPtr header, int size);

        [DllImport("winmm.dll")]
        private static extern int waveOutUnprepareHeader(IntPtr hwo, IntPtr header, int size);

        [DllImport("winmm.dll")]
        private static extern int waveOutWrite(IntPtr hwo, IntPtr header, int size);

        [DllImport("winmm.dll")]
        private static extern int waveOutReset(IntPtr hwo);

        [DllImport("winmm.dll")]
        private static extern int waveOutClose(IntPtr hwo);

        // kept in a static field: the callback must not be collected while a sound is playing
        private static readonly WaveCallback Callback = OnWaveMessage;

        private class Voice
        {
            public long Id;
            public IntPtr Handle;
            public IntPtr Header;
            public IntPtr Data;
            public bool Done;
        }

        // Plays the sound of one motion (WobbleType.None and WobbleType.Random have none). Nothing here may throw: a missing sound device is not an error.
        public static void Play(WobbleType type, int volume)
        {
            try
            {
                short[] pcm = Load(type);
                if (pcm == null || volume <= 0) return;
                lock (Lock)
                {
                    Reap();
                    // the older voices keep their level; the new one is made a little quieter for each sound that is still playing
                    int playing = Voices.Count;
                    while (Voices.Count >= MaxVoices) StopOldest();
                    double gain = Math.Pow(Math.Max(0, Math.Min(100, volume)) / 100.0, 2) / Math.Sqrt(playing + 1);
                    byte[] bytes = Render(pcm, gain);
                    Start(bytes);
                }
            }
            catch (Exception)
            {
            }
        }

        // The samples with the volume, the fade in at the start and the fade out at the end (a cosine curve), and a little silence after them.
        private static byte[] Render(short[] pcm, double gain)
        {
            int tail = SampleRate * TailMs / 1000;
            int fadeIn = Math.Min(pcm.Length, SampleRate * FadeInMs / 1000);
            int fadeOut = Math.Min(pcm.Length, SampleRate * FadeOutMs / 1000);
            var result = new byte[(pcm.Length + tail) * 2];
            for (int i = 0; i < pcm.Length; i++)
            {
                double v = pcm[i] * gain;
                if (i < fadeIn) v *= 0.5 * (1 - Math.Cos(Math.PI * i / fadeIn));
                int fromEnd = pcm.Length - 1 - i;
                if (fromEnd < fadeOut) v *= 0.5 * (1 + Math.Cos(Math.PI * (fadeOut - fromEnd) / fadeOut));
                short s = (short)Math.Max(short.MinValue, Math.Min(short.MaxValue, Math.Round(v)));
                result[i * 2] = (byte)(s & 0xFF);
                result[i * 2 + 1] = (byte)((s >> 8) & 0xFF);
            }
            return result;
        }

        private static void Start(byte[] bytes)
        {
            var format = new WaveFormat
            {
                wFormatTag = 1,
                nChannels = 1,
                nSamplesPerSec = SampleRate,
                wBitsPerSample = 16,
                nBlockAlign = 2,
                nAvgBytesPerSec = SampleRate * 2,
                cbSize = 0
            };
            IntPtr hwo;
            // the id of the voice goes to the callback: a handle may be used again after it has been closed, an id never
            long id = ++_nextId;
            if (waveOutOpen(out hwo, WaveMapper, ref format, Callback, new IntPtr(id), CallbackFunction) != 0) return;
            var voice = new Voice { Id = id, Handle = hwo };
            int headerSize = Marshal.SizeOf(typeof(WaveHeader));
            voice.Data = Marshal.AllocHGlobal(bytes.Length);
            Marshal.Copy(bytes, 0, voice.Data, bytes.Length);
            voice.Header = Marshal.AllocHGlobal(headerSize);
            var header = new WaveHeader { lpData = voice.Data, dwBufferLength = (uint)bytes.Length };
            Marshal.StructureToPtr(header, voice.Header, false);
            if (waveOutPrepareHeader(hwo, voice.Header, headerSize) != 0 || waveOutWrite(hwo, voice.Header, headerSize) != 0)
            {
                Release(voice);
                return;
            }
            Voices.Add(voice);
        }

        // The device tells that a sound is over. Nothing of the waveOut API may be called in this callback, so the voice is only marked.
        private static void OnWaveMessage(IntPtr hwo, uint message, IntPtr instance, IntPtr param1, IntPtr param2)
        {
            if (message != WomDone) return;
            // no lock here: a thread that holds it may be waiting in waveOutReset or waveOutClose for this very message
            DoneIds.Enqueue(instance.ToInt64());
            ThreadPool.QueueUserWorkItem(_ => { try { lock (Lock) Reap(); } catch (Exception) { } });
        }

        // Closes the voices that have finished.
        private static void Reap()
        {
            long finished;
            while (DoneIds.TryDequeue(out finished))
                foreach (Voice v in Voices) if (v.Id == finished) v.Done = true;
            for (int i = Voices.Count - 1; i >= 0; i--)
            {
                if (!Voices[i].Done) continue;
                Release(Voices[i]);
                Voices.RemoveAt(i);
            }
        }

        private static void StopOldest()
        {
            Voice oldest = Voices[0];
            Voices.RemoveAt(0);
            try { waveOutReset(oldest.Handle); } catch (Exception) { }
            Release(oldest);
        }

        private static void Release(Voice voice)
        {
            try
            {
                int headerSize = Marshal.SizeOf(typeof(WaveHeader));
                waveOutReset(voice.Handle); // makes sure that the device no longer reads the buffer that is freed below
                if (voice.Header != IntPtr.Zero) waveOutUnprepareHeader(voice.Handle, voice.Header, headerSize);
                waveOutClose(voice.Handle);
            }
            catch (Exception)
            {
            }
            if (voice.Header != IntPtr.Zero) Marshal.FreeHGlobal(voice.Header);
            if (voice.Data != IntPtr.Zero) Marshal.FreeHGlobal(voice.Data);
            voice.Header = voice.Data = IntPtr.Zero;
        }

        // The samples of one embedded wave file (16 bit, after the 44 byte header).
        private static short[] Load(WobbleType type)
        {
            lock (Lock)
            {
                short[] data;
                if (Cache.TryGetValue(type, out data)) return data;
                string name;
                if (!Resources.TryGetValue(type, out name)) return null;
                byte[] bytes;
                using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("VerticalSheetTabs.Sounds." + name + ".wav"))
                {
                    if (s == null) return null;
                    bytes = new byte[s.Length];
                    int read = 0;
                    while (read < bytes.Length)
                    {
                        int n = s.Read(bytes, read, bytes.Length - read);
                        if (n <= 0) break;
                        read += n;
                    }
                }
                int count = Math.Max(0, (bytes.Length - 44) / 2);
                data = new short[count];
                for (int i = 0; i < count; i++) data[i] = (short)(bytes[44 + i * 2] | (bytes[45 + i * 2] << 8));
                Cache[type] = data;
                return data;
            }
        }
    }
}
