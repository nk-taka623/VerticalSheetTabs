using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace VerticalSheetTabs
{
    // The preview of the wallpaper picture in the settings dialog, with the cropping frame on it. The frame is held by its corners
    // and edges (it changes its size) and by its inside (it moves); what is outside it is darkened, because it will not be shown.
    // The arithmetic is in WallpaperCrop.cs (CropDrag, CropPreviewMath); this control draws and passes the mouse on.
    public class CropPreview : Control
    {
        private Bitmap _image;          // the picture (this control keeps it, and disposes it)
        private Bitmap _scaled;         // the picture in the size it is drawn in
        private WallpaperCrop _crop = WallpaperCrop.Full;
        private string _message = "";
        private CropHandle _hold = CropHandle.None;
        private RectangleF _holdFrame;
        private PointF _holdPoint;

        // The frame changed (the mouse moved it, or the reset button did).
        public event EventHandler CropChanged;

        public CropPreview()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.Selectable, false);
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            TabStop = false;
            BackColor = SystemColors.ControlLight;
            MinimumSize = new Size(120, 90);
        }

        public bool HasImage { get { return _image != null; } }

        public Size ImageSize { get { return _image == null ? Size.Empty : _image.Size; } }

        public WallpaperCrop Crop
        {
            get { return _crop; }
            set
            {
                WallpaperCrop n = value.Normalize();
                if (_crop.SameAs(n)) return;
                _crop = n;
                Invalidate();
            }
        }

        // What is written where there is no picture (none chosen, not found, cannot be read).
        public string Message
        {
            get { return _message; }
            set { _message = value ?? ""; if (_image == null) Invalidate(); }
        }

        // Takes the picture (null: no picture) and the frame. The picture it had is disposed.
        public void SetImage(Bitmap image, WallpaperCrop crop)
        {
            EndHold();
            if (_scaled != null) { _scaled.Dispose(); _scaled = null; }
            if (_image != null && !ReferenceEquals(_image, image)) _image.Dispose();
            _image = image;
            _crop = image == null ? WallpaperCrop.Full : crop.Normalize();
            Cursor = Cursors.Default;
            Invalidate();
        }

        // Puts the frame back to the whole picture.
        public void ResetCrop()
        {
            if (_crop.IsFull) return;
            _crop = WallpaperCrop.Full;
            Invalidate();
            RaiseChanged();
        }

        private void RaiseChanged()
        {
            EventHandler handler = CropChanged;
            if (handler != null) handler(this, EventArgs.Empty);
        }

        private int Dip(int v) { return Math.Max(1, (int)Math.Round(v * DeviceDpi / 96.0)); }

        // Where the picture lies in the control (it is as large as it can be, with a margin for the handles).
        private RectangleF PictureBounds()
        {
            if (_image == null) return RectangleF.Empty;
            return CropPreviewMath.FitRect(ClientSize.Width, ClientSize.Height, _image.Width, _image.Height, Dip(12));
        }

        private RectangleF Frame(RectangleF bounds)
        {
            return CropPreviewMath.ToFrame(_crop, bounds);
        }

        private void EndHold()
        {
            _hold = CropHandle.None;
            if (Capture) Capture = false;
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            try
            {
                if (e.Button != MouseButtons.Left || _image == null) return;
                RectangleF bounds = PictureBounds();
                if (bounds.IsEmpty) return;
                var p = new PointF(e.X, e.Y);
                RectangleF frame = Frame(bounds);
                CropHandle h = CropDrag.HitTest(frame, p, Dip(8), !_crop.IsFull);
                if (h == CropHandle.None) return;
                _hold = h;
                _holdFrame = frame;
                _holdPoint = p;
                Capture = true;
            }
            catch (Exception)
            {
                EndHold();
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            try
            {
                if (_image == null) return;
                RectangleF bounds = PictureBounds();
                if (bounds.IsEmpty) return;
                if (_hold != CropHandle.None)
                {
                    RectangleF frame = CropDrag.Drag(_hold, _holdFrame, _holdPoint, new PointF(e.X, e.Y), bounds, Dip(24), Dip(24));
                    WallpaperCrop crop = CropPreviewMath.FromFrame(frame, bounds);
                    if (!crop.SameAs(_crop))
                    {
                        _crop = crop;
                        Invalidate();
                        RaiseChanged();
                    }
                    return;
                }
                Cursor = CursorFor(CropDrag.HitTest(Frame(bounds), new PointF(e.X, e.Y), Dip(8), !_crop.IsFull));
            }
            catch (Exception)
            {
                EndHold();
            }
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (e.Button == MouseButtons.Left) EndHold();
        }

        protected override void OnMouseCaptureChanged(EventArgs e)
        {
            base.OnMouseCaptureChanged(e);
            if (!Capture) _hold = CropHandle.None;
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            if (_hold == CropHandle.None) Cursor = Cursors.Default;
        }

        private static Cursor CursorFor(CropHandle h)
        {
            switch (h)
            {
                case CropHandle.NW:
                case CropHandle.SE: return Cursors.SizeNWSE;
                case CropHandle.NE:
                case CropHandle.SW: return Cursors.SizeNESW;
                case CropHandle.N:
                case CropHandle.S: return Cursors.SizeNS;
                case CropHandle.W:
                case CropHandle.E: return Cursors.SizeWE;
                case CropHandle.Move: return Cursors.SizeAll;
                default: return Cursors.Default;
            }
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (_scaled != null) { _scaled.Dispose(); _scaled = null; }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            try
            {
                Graphics g = e.Graphics;
                g.Clear(BackColor);
                if (_image == null)
                {
                    if (_message.Length > 0)
                        TextRenderer.DrawText(g, _message, Font, ClientRectangle, SystemColors.GrayText, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.WordBreak);
                    return;
                }
                RectangleF bounds = PictureBounds();
                if (bounds.IsEmpty) return;
                Rectangle pic = Rectangle.Round(bounds);
                if (_scaled == null || _scaled.Width != pic.Width || _scaled.Height != pic.Height)
                {
                    if (_scaled != null) _scaled.Dispose();
                    _scaled = new Bitmap(Math.Max(1, pic.Width), Math.Max(1, pic.Height), System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
                    using (Graphics sg = Graphics.FromImage(_scaled))
                    {
                        sg.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        sg.PixelOffsetMode = PixelOffsetMode.HighQuality;
                        sg.DrawImage(_image, new Rectangle(0, 0, _scaled.Width, _scaled.Height), 0, 0, _image.Width, _image.Height, GraphicsUnit.Pixel);
                    }
                }
                g.DrawImageUnscaled(_scaled, pic.Location);

                // what is outside the frame is not shown: darken it
                Rectangle fr = Rectangle.Round(Frame(bounds));
                using (var shade = new SolidBrush(Color.FromArgb(150, 0, 0, 0)))
                {
                    g.FillRectangle(shade, pic.Left, pic.Top, pic.Width, Math.Max(0, fr.Top - pic.Top));
                    g.FillRectangle(shade, pic.Left, fr.Bottom, pic.Width, Math.Max(0, pic.Bottom - fr.Bottom));
                    g.FillRectangle(shade, pic.Left, fr.Top, Math.Max(0, fr.Left - pic.Left), fr.Height);
                    g.FillRectangle(shade, fr.Right, fr.Top, Math.Max(0, pic.Right - fr.Right), fr.Height);
                }

                // the frame: a dark line outside and a white one inside, so that it shows on any picture
                using (var dark = new Pen(Color.FromArgb(200, 0, 0, 0), 1f))
                using (var white = new Pen(Color.White, 1f))
                {
                    g.DrawRectangle(dark, fr.Left - 1, fr.Top - 1, fr.Width + 2, fr.Height + 2);
                    g.DrawRectangle(white, fr.Left, fr.Top, fr.Width, fr.Height);

                    // the middle of the frame: dotted lines, white over dark
                    int cx = fr.Left + fr.Width / 2, cy = fr.Top + fr.Height / 2;
                    using (var d1 = new Pen(Color.FromArgb(220, 0, 0, 0), 1f) { DashPattern = new float[] { 3f, 3f } })
                    using (var d2 = new Pen(Color.White, 1f) { DashPattern = new float[] { 3f, 3f }, DashOffset = 3f })
                    {
                        g.DrawLine(d1, cx, fr.Top + 1, cx, fr.Bottom - 1);
                        g.DrawLine(d2, cx, fr.Top + 1, cx, fr.Bottom - 1);
                        g.DrawLine(d1, fr.Left + 1, cy, fr.Right - 1, cy);
                        g.DrawLine(d2, fr.Left + 1, cy, fr.Right - 1, cy);
                    }

                    // the corners: small white squares
                    int hs = Dip(9), half = hs / 2;
                    Point[] corners = { new Point(fr.Left, fr.Top), new Point(fr.Right, fr.Top), new Point(fr.Left, fr.Bottom), new Point(fr.Right, fr.Bottom) };
                    foreach (Point c in corners)
                    {
                        var r = new Rectangle(c.X - half, c.Y - half, hs, hs);
                        g.FillRectangle(Brushes.White, r);
                        g.DrawRectangle(dark, r);
                    }
                }
            }
            catch (Exception)
            {
                // A preview that cannot be drawn is left plain: an exception here would break the whole dialog.
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_scaled != null) { _scaled.Dispose(); _scaled = null; }
                if (_image != null) { _image.Dispose(); _image = null; }
            }
            base.Dispose(disposing);
        }
    }
}
