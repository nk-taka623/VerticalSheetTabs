using System;

namespace VerticalSheetTabs
{
    // The wobble of the selected sheet. 1 to 6 are the motions, 7 picks one of them at each selection.
    public enum WobbleType
    {
        None = 0,
        Shake = 1,  // sideways
        Wobble = 2, // rotating around the center
        Jelly = 3,  // squash and stretch
        Bounce = 4, // hopping
        Pulse = 5,  // growing and shrinking
        Swing = 6,  // a pendulum around the upper left corner
        Random = 7
    }

    // One frame of a motion: a move (DIP), a turn (degrees), a scale, and the pivot of the turn and the scale.
    public struct WobbleFrame
    {
        public double Dx, Dy, Angle, ScaleX, ScaleY;
        public bool PivotTopLeft;
    }

    public static class WobbleMotion
    {
        public const int AmpMin = 40, AmpMax = 200, AmpDefault = 100;     // percent
        public const int DurMin = 300, DurMax = 1200, DurDefault = 600;   // milliseconds

        private static readonly Random Rng = new Random();

        public static WobbleType PickRandom()
        {
            return (WobbleType)Rng.Next(1, 7);
        }

        // t runs from 0 to 1; every motion dies down to nothing at t = 1. amp is 1.0 for 100%.
        public static WobbleFrame At(WobbleType type, double t, double amp)
        {
            t = Math.Max(0, Math.Min(1, t));
            double decay = (1 - t) * (1 - t); // eased out: 1 at the start, 0 at the end
            var f = new WobbleFrame { ScaleX = 1, ScaleY = 1 };
            switch (type)
            {
                case WobbleType.Shake:
                    f.Dx = amp * 6 * decay * Math.Sin(2 * Math.PI * 3 * t);
                    break;
                case WobbleType.Wobble:
                    f.Angle = amp * 3 * decay * Math.Sin(2 * Math.PI * 2.5 * t);
                    break;
                case WobbleType.Jelly:
                    {
                        double w = decay * Math.Sin(2 * Math.PI * 2.5 * t + Math.PI);
                        f.ScaleX = 1 - amp * 0.06 * w;
                        f.ScaleY = 1 + amp * 0.12 * w;
                        break;
                    }
                case WobbleType.Bounce:
                    f.Dy = -amp * 8 * decay * Math.Abs(Math.Sin(2 * Math.PI * 1.5 * t));
                    break;
                case WobbleType.Pulse:
                    f.ScaleX = f.ScaleY = 1 + amp * 0.07 * decay * Math.Abs(Math.Sin(2 * Math.PI * 1.5 * t));
                    break;
                case WobbleType.Swing:
                    f.Angle = amp * 4 * decay * Math.Sin(2 * Math.PI * 2 * t);
                    f.PivotTopLeft = true;
                    break;
            }
            return f;
        }
    }
}
