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
        private WallpaperRotation _rotation;    // how the picture is turned (not rounded while the pointer drags)
        private bool _hoverRotate;              // the pointer is near the dotted lines: the circle is shown
        private bool _rotateBlocked;            // the turn the pointer asks for is not taken (the frame would be too small)
        private PointF _rotatePivot;            // where the pointer goes round: the middle of the frame, and once the frame is brought to the middle of the preview, that
        private bool _recentered;               // the frame was brought to the middle of the preview (when the turn began)
        private double _rotateAngle;            // the angle of the pointer at the last turn that was taken
        private bool _rotateAngleKnown;
        private double _pointerAngle;
        private bool _limit = true;             // the frame is kept inside the (turned) picture; off: it may reach outside, anywhere in the area
        private string _message = "";
        private CropHandle _hold = CropHandle.None;
        private RectangleF _holdFrame;
        private PointF _holdPoint;

        // The frame or the turn changed (the mouse moved it, or the reset button did).
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

        public WallpaperRotation Rotation
        {
            get
            {
                RectangleF bounds = PictureBounds();
                return bounds.IsEmpty ? _rotation.Normalize() : Live(bounds).Normalize();
            }
            set
            {
                WallpaperRotation n = value.Normalize();
                if (_rotation.SameAs(n)) return;
                _rotation = n;
                Invalidate();
            }
        }

        // The frame is kept inside the picture ("limit the frame" on). Turning it on brings a frame that reaches outside back in.
        public bool LimitCrop
        {
            get { return _limit; }
            set
            {
                if (_limit == value) return;
                _limit = value;
                if (value && _image != null) FitToPicture();
                Invalidate();
            }
        }

        // What is written where there is no picture (none chosen, not found, cannot be read).
        public string Message
        {
            get { return _message; }
            set { _message = value ?? ""; if (_image == null) Invalidate(); }
        }

        // Takes the picture (null: no picture), the frame and the turn. The picture it had is disposed.
        public void SetImage(Bitmap image, WallpaperCrop crop, WallpaperRotation rotation, bool limit = true)
        {
            EndHold();
            _limit = limit;
            if (_scaled != null) { _scaled.Dispose(); _scaled = null; }
            if (_image != null && !ReferenceEquals(_image, image)) _image.Dispose();
            _image = image;
            _crop = image == null ? WallpaperCrop.Full : crop.Normalize();
            _rotation = image == null ? WallpaperRotation.None : rotation.Normalize();
            _hoverRotate = false;
            Cursor = Cursors.Default;
            Invalidate();
        }

        // Puts the frame back to the whole picture, and the picture back to not turned.
        public void ResetCrop()
        {
            if (_crop.IsFull && _rotation.IsIdentity && _limit) return;
            _limit = true;
            _crop = WallpaperCrop.Full;
            _rotation = WallpaperRotation.None;
            Invalidate();
            RaiseChanged();
        }

        // The frame is made smaller (round its middle, with its ratio) until it lies inside the picture as it is shown (turned, and
        // made smaller by the zoom) and inside the place of the picture; a frame whose middle is outside is first brought to the
        // middle of the preview.
        private void FitToPicture()
        {
            RectangleF bounds = PictureBounds();
            if (bounds.IsEmpty) return;
            RectangleF frame = Frame(bounds);
            PointF[] corners = RotationMath.Corners(Live(bounds), bounds);
            double inset = Inset(bounds);
            PointF mid = RotationMath.Center(bounds);
            if (!RotationMath.Inside(corners, new PointF(frame.Left + frame.Width / 2f, frame.Top + frame.Height / 2f), inset))
                frame = new RectangleF(mid.X - frame.Width / 2f, mid.Y - frame.Height / 2f, frame.Width, frame.Height);
            RectangleF fitted;
            if (!RotationMath.FitFrame(frame, corners, inset, out fitted)) fitted = frame;
            RectangleF plain = fitted;
            if (RotationMath.FitFrame(fitted, RotationMath.Corners(WallpaperRotation.None, bounds), inset, out plain)) fitted = plain;
            WallpaperCrop crop = CropPreviewMath.FromFrame(fitted, bounds);
            if (crop.SameAs(_crop)) return;
            _crop = crop;
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

        // The largest area the preview can show (the whole control, less the margin): the turned picture is kept inside it.
        private RectangleF PictureArea()
        {
            int m = Dip(12);
            return new RectangleF(m, m, Math.Max(1, ClientSize.Width - 2 * m), Math.Max(1, ClientSize.Height - 2 * m));
        }

        private double ZoomFor(double degrees, RectangleF bounds)
        {
            RectangleF area = PictureArea();
            return RotationMath.FitZoom(degrees, bounds.Width, bounds.Height, area.Width, area.Height);
        }

        // The turn as it is shown now: the angle, and the zoom that follows from it and the size of the preview.
        private WallpaperRotation Live(RectangleF bounds)
        {
            return new WallpaperRotation(_rotation.Degrees, ZoomFor(_rotation.Degrees, bounds));
        }

        // How near to a side of the turned picture a corner of the frame may come: two pixels of the picture, at least one of the preview.
        private double Inset(RectangleF bounds)
        {
            return _image == null ? 1.0 : Math.Max(1.0, 2.0 * bounds.Width / _image.Width);
        }

        private RectangleF Frame(RectangleF bounds)
        {
            return CropPreviewMath.ToFrame(_crop, bounds);
        }

        private void EndHold()
        {
            bool turning = _hold == CropHandle.Rotate;
            _hold = CropHandle.None;
            _rotateBlocked = false;
            _rotateAngleKnown = false;
            _recentered = false;
            if (turning) Invalidate();
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
                if (h == CropHandle.Rotate)
                {
                    _rotatePivot = new PointF(frame.Left + frame.Width / 2f, frame.Top + frame.Height / 2f);
                    _recentered = false;
                    _rotateBlocked = false;
                    _rotateAngleKnown = false;
                    StartTurn(p);
                }
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
                if (_hold == CropHandle.Rotate)
                {
                    Turn(new PointF(e.X, e.Y), bounds);
                    return;
                }
                if (_hold != CropHandle.None)
                {
                    RectangleF frame = CropDrag.Drag(_hold, _holdFrame, _holdPoint, new PointF(e.X, e.Y), _limit ? bounds : PictureArea(), Dip(24), Dip(24));
                    if (_limit && !Live(bounds).IsIdentity)
                        frame = RotationMath.LimitFrame(Frame(bounds), frame, RotationMath.Corners(Live(bounds), bounds), Inset(bounds), _hold == CropHandle.Move);
                    WallpaperCrop crop = CropPreviewMath.FromFrame(frame, bounds);
                    if (!crop.SameAs(_crop))
                    {
                        _crop = crop;
                        Invalidate();
                        RaiseChanged();
                    }
                    return;
                }
                CropHandle over = CropDrag.HitTest(Frame(bounds), new PointF(e.X, e.Y), Dip(8), !_crop.IsFull);
                Cursor = CursorFor(over);
                bool near = over == CropHandle.Rotate;
                if (near != _hoverRotate)
                {
                    _hoverRotate = near;
                    Invalidate();
                }
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
            if (_hold == CropHandle.None)
            {
                Cursor = Cursors.Default;
                if (_hoverRotate)
                {
                    _hoverRotate = false;
                    Invalidate();
                }
            }
        }

        // The pointer is put on the dotted lines to turn the picture: the angle it is at now is where the turn goes on from.
        private void StartTurn(PointF p)
        {
            double dx = p.X - _rotatePivot.X, dy = p.Y - _rotatePivot.Y;
            if (dx * dx + dy * dy < 9) return;
            _rotateAngle = _pointerAngle = RotationMath.AngleOf(_rotatePivot, p);
            _rotateAngleKnown = true;
        }

        // The pointer moved while the picture is turned: the picture is turned by the angle the pointer went round the middle of the
        // frame; the frame gets smaller (round its middle) if it would stick out of the turned picture. A turn that would make
        // the frame too small is not taken.
        private void Turn(PointF p, RectangleF bounds)
        {
            double dx = p.X - _rotatePivot.X, dy = p.Y - _rotatePivot.Y;
            if (dx * dx + dy * dy < 9) return;
            double angle = RotationMath.AngleOf(_rotatePivot, p);
            _pointerAngle = angle;
            if (!_rotateAngleKnown)
            {
                _rotateAngle = angle;
                _rotateAngleKnown = true;
                Invalidate();
                return;
            }
            double delta = RotationMath.Delta(_rotateAngle, angle);
            if (Math.Abs(delta) < 1e-6)
            {
                Invalidate();
                return;
            }
            // the picture is turned round the middle of the preview; when the turn begins, the frame is brought to that middle too
            double deg = WallpaperRotation.WrapDegrees(_rotation.Degrees + delta);
            var next = new WallpaperRotation(deg, ZoomFor(deg, bounds));
            RectangleF start = Frame(bounds);
            PointF mid = RotationMath.Center(bounds);
            if (!_recentered) start = new RectangleF(mid.X - start.Width / 2f, mid.Y - start.Height / 2f, start.Width, start.Height);
            RectangleF fitted;
            float minW = Math.Min(Dip(24), bounds.Width / 2f), minH = Math.Min(Dip(24), bounds.Height / 2f);
            // (the frame is made smaller only when it is kept inside the picture; otherwise it stays as it is)
            bool fits = _limit ? RotationMath.FitFrame(start, RotationMath.Corners(next, bounds), Inset(bounds), out fitted) : (fitted = start).Width > 0;
            if (fits && fitted.Width >= minW && fitted.Height >= minH)
            {
                _rotation = next;
                _crop = CropPreviewMath.FromFrame(fitted, bounds);
                if (!_recentered)
                {
                    // the circle is round the middle of the frame, which has moved: the pointer goes round the new middle from here
                    _recentered = true;
                    _rotatePivot = mid;
                    _rotateAngle = _pointerAngle = RotationMath.AngleOf(mid, p);
                }
                else
                {
                    _rotateAngle = angle;
                }
                _rotateBlocked = false;
                Invalidate();
                RaiseChanged();
            }
            else
            {
                _rotateBlocked = true;
                Invalidate();
            }
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
                case CropHandle.Rotate: return Cursors.Hand;
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
                WallpaperRotation live = Live(bounds);
                RectangleF shaded = bounds;
                if (live.IsIdentity && _limit)
                {
                    g.DrawImageUnscaled(_scaled, pic.Location);
                }
                else if (live.IsIdentity)
                {
                    // (the frame may reach outside the picture: the whole area is shown, grey where there is no picture)
                    shaded = PictureArea();
                    using (var grey = new SolidBrush(Wallpaper.NoPictureGrey)) g.FillRectangle(grey, shaded);
                    g.DrawImageUnscaled(_scaled, pic.Location);
                }
                else
                {
                    // the turned picture (made smaller if it would stick out of the area), cut off at the edge of the area; where
                    // it does not reach inside the preview, a dark grey
                    RectangleF area = PictureArea();
                    shaded = area;
                    GraphicsState saved = g.Save();
                    using (var grey = new SolidBrush(Wallpaper.NoPictureGrey)) g.FillRectangle(grey, _limit ? bounds : area);
                    g.SetClip(area);
                    using (var m = new Matrix())
                    {
                        PointF mid = RotationMath.Center(bounds);
                        m.Translate(mid.X, mid.Y);
                        m.Rotate((float)live.Degrees);
                        m.Scale((float)live.Zoom, (float)live.Zoom);
                        m.Translate(-mid.X, -mid.Y);
                        g.Transform = m;
                        g.InterpolationMode = InterpolationMode.HighQualityBilinear;
                        g.PixelOffsetMode = PixelOffsetMode.Half;
                        g.DrawImage(_scaled, bounds, new RectangleF(0, 0, _scaled.Width, _scaled.Height), GraphicsUnit.Pixel);
                    }
                    g.Restore(saved);
                }

                // what is outside the frame is not shown: darken it
                Rectangle fr = Rectangle.Round(Frame(bounds));
                Rectangle sh = Rectangle.Round(shaded);
                using (var shade = new SolidBrush(Color.FromArgb(150, 0, 0, 0)))
                {
                    g.FillRectangle(shade, sh.Left, sh.Top, sh.Width, Math.Max(0, fr.Top - sh.Top));
                    g.FillRectangle(shade, sh.Left, fr.Bottom, sh.Width, Math.Max(0, sh.Bottom - fr.Bottom));
                    g.FillRectangle(shade, sh.Left, fr.Top, Math.Max(0, fr.Left - sh.Left), fr.Height);
                    g.FillRectangle(shade, fr.Right, fr.Top, Math.Max(0, sh.Right - fr.Right), fr.Height);
                }
                // the frame: a dark line outside and a white one inside, so that it shows on any picture
                using (var dark = new Pen(Color.FromArgb(200, 0, 0, 0), 1f))
                using (var white = new Pen(Color.White, 1f))
                {
                    g.DrawRectangle(dark, fr.Left - 1, fr.Top - 1, fr.Width + 2, fr.Height + 2);
                    g.DrawRectangle(white, fr.Left, fr.Top, fr.Width, fr.Height);

                    // the middle of the frame: dotted lines, white over dark
                    int cx = fr.Left + fr.Width / 2, cy = fr.Top + fr.Height / 2;
                    bool circle = _hoverRotate || _hold == CropHandle.Rotate;
                    Color lineColor = circle ? (_rotateBlocked ? Color.FromArgb(255, 70, 60) : Color.FromArgb(255, 220, 60)) : Color.White;
                    float lineWidth = circle ? 2f : 1f;
                    using (var d1 = new Pen(Color.FromArgb(220, 0, 0, 0), lineWidth) { DashPattern = new float[] { 3f, 3f } })
                    using (var d2 = new Pen(lineColor, lineWidth) { DashPattern = new float[] { 3f, 3f }, DashOffset = 3f })
                    {
                        g.DrawLine(d1, cx, fr.Top + 1, cx, fr.Bottom - 1);
                        g.DrawLine(d2, cx, fr.Top + 1, cx, fr.Bottom - 1);
                        g.DrawLine(d1, fr.Left + 1, cy, fr.Right - 1, cy);
                        g.DrawLine(d2, fr.Left + 1, cy, fr.Right - 1, cy);
                    }

                    if (circle) DrawTurnCircle(g, fr, cx, cy, lineColor);

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

        // The circle round the crossing of the dotted lines: it shows where the picture is turned from (a quarter of the short side
        // of the frame is its radius). While it is turned, a mark on the circle follows the pointer and the angle is written.
        private void DrawTurnCircle(Graphics g, Rectangle fr, int cx, int cy, Color color)
        {
            float r = Math.Min(fr.Width, fr.Height) / 4f;
            if (r < 2f) return;
            SmoothingMode oldMode = g.SmoothingMode;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var shadow = new Pen(Color.FromArgb(180, 0, 0, 0), 5f))
            using (var ring = new Pen(color, 3f))
            {
                g.DrawEllipse(shadow, cx - r, cy - r, 2 * r, 2 * r);
                g.DrawEllipse(ring, cx - r, cy - r, 2 * r, 2 * r);
            }
            using (var dot = new SolidBrush(color)) g.FillEllipse(dot, cx - 3f, cy - 3f, 6f, 6f);
            if (_hold == CropHandle.Rotate && _rotateAngleKnown)
            {
                double a = _pointerAngle * Math.PI / 180.0;
                float px = cx + r * (float)Math.Cos(a), py = cy + r * (float)Math.Sin(a);
                using (var mark = new SolidBrush(Color.FromArgb(255, 90, 40)))
                {
                    g.FillEllipse(mark, px - 6f, py - 6f, 12f, 12f);
                    g.DrawEllipse(Pens.White, px - 6f, py - 6f, 12f, 12f);
                }
                string text = _rotation.Normalize().Degrees.ToString("0.#", System.Globalization.CultureInfo.CurrentCulture) + "°";
                SizeF size = g.MeasureString(text, Font);
                float tx = cx + r + 8f, ty = cy - size.Height / 2f;
                if (tx + size.Width > fr.Right) tx = cx - size.Width / 2f;
                using (var back = new SolidBrush(Color.FromArgb(170, 0, 0, 0))) g.FillRectangle(back, tx - 2f, ty - 1f, size.Width + 4f, size.Height + 2f);
                g.DrawString(text, Font, Brushes.White, tx, ty);
            }
            g.SmoothingMode = oldMode;
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
