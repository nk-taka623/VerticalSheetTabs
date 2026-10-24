using System;
using System.Collections.Generic;
using System.IO;
using System.Media;
using System.Reflection;

namespace VerticalSheetTabs
{
    // The sound effects of the wobble animations (Agents.md 12.172): one short sound for each motion, embedded in the
    // add-in as 16 bit mono wave files (Sounds\*.wav). The volume is applied to the samples before they are played.
    public static class WobbleSounds
    {
        public const int VolumeDefault = 60;

        private static readonly Dictionary<WobbleType, string> Resources = new Dictionary<WobbleType, string>
        {
            { WobbleType.Shake, "shake" },
            { WobbleType.Wobble, "wobble" },
            { WobbleType.Jelly, "jelly" },
            { WobbleType.Bounce, "bounce" },
            { WobbleType.Pulse, "pulse" },
            { WobbleType.Swing, "swing" }
        };

        private static readonly Dictionary<WobbleType, byte[]> Cache = new Dictionary<WobbleType, byte[]>();
        private static SoundPlayer _player;
        private static MemoryStream _stream;

        // Plays the sound of one motion (WobbleType.None and WobbleType.Random have none). Nothing here may throw: a missing sound device is not an error.
        public static void Play(WobbleType type, int volume)
        {
            try
            {
                byte[] wave = Load(type);
                if (wave == null || volume <= 0) return;
                byte[] scaled = Scale(wave, volume);
                Stop();
                _stream = new MemoryStream(scaled);
                _player = new SoundPlayer(_stream);
                _player.Play();
            }
            catch (Exception)
            {
            }
        }

        private static void Stop()
        {
            try
            {
                if (_player != null) { _player.Stop(); _player.Dispose(); }
                if (_stream != null) _stream.Dispose();
            }
            catch (Exception)
            {
            }
            _player = null;
            _stream = null;
        }

        private static byte[] Load(WobbleType type)
        {
            byte[] data;
            if (Cache.TryGetValue(type, out data)) return data;
            string name;
            if (!Resources.TryGetValue(type, out name)) return null;
            using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("VerticalSheetTabs.Sounds." + name + ".wav"))
            {
                if (s == null) return null;
                data = new byte[s.Length];
                int read = 0;
                while (read < data.Length)
                {
                    int n = s.Read(data, read, data.Length - read);
                    if (n <= 0) break;
                    read += n;
                }
            }
            Cache[type] = data;
            return data;
        }

        // A copy of the wave with the samples (16 bit, after the 44 byte header) scaled: the volume 0 to 100 is applied squared, which sounds more even.
        private static byte[] Scale(byte[] wave, int volume)
        {
            double gain = Math.Pow(Math.Max(0, Math.Min(100, volume)) / 100.0, 2);
            var copy = (byte[])wave.Clone();
            for (int i = 44; i + 1 < copy.Length; i += 2)
            {
                short v = (short)(copy[i] | (copy[i + 1] << 8));
                v = (short)Math.Round(v * gain);
                copy[i] = (byte)(v & 0xFF);
                copy[i + 1] = (byte)((v >> 8) & 0xFF);
            }
            return copy;
        }
    }
}
