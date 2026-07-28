using System.Drawing.Drawing2D;

namespace EmailReviewViewer.App;

public sealed class RoundedInputHost : Panel
{
    private readonly Control _input;
    private readonly Panel? _nativeClip;
    private int _cornerRadius = AppTheme.CornerRadius;
    private int _borderThickness = AppTheme.BorderThickness;

    public int CornerRadius
    {
        get => _cornerRadius;
        set
        {
            _cornerRadius = Math.Max(0, value);
            UpdateHostRegion();
            Invalidate();
        }
    }

    public int BorderThickness
    {
        get => _borderThickness;
        set
        {
            _borderThickness = Math.Max(1, value);
            PerformLayout();
            Invalidate();
        }
    }

    public RoundedInputHost(Control input)
    {
        _input = input;
        var width = Math.Max(80, input.Width);
        var margin = input.Margin;

        SetStyle(
            ControlStyles.UserPaint |
            ControlStyles.AllPaintingInWmPaint |
            ControlStyles.OptimizedDoubleBuffer |
            ControlStyles.ResizeRedraw,
            true);

        BackColor = AppTheme.Current.Surface;
        Cursor = input is TextBox ? Cursors.IBeam : Cursors.Default;
        Height = AppTheme.ControlHeight;
        MinimumSize = new Size(0, AppTheme.ControlHeight);
        Width = width;
        Margin = margin;
        TabStop = false;

        if (input is TextBox textBox)
            textBox.BorderStyle = BorderStyle.None;
        input.BackColor = AppTheme.Current.Surface;
        input.ForeColor = AppTheme.Current.Text;
        input.Margin = Padding.Empty;
        input.MinimumSize = Size.Empty;
        input.Dock = DockStyle.None;
        input.GotFocus += (_, _) => Invalidate();
        input.LostFocus += (_, _) => Invalidate();
        input.EnabledChanged += (_, _) => Invalidate();
        if (input is DateTimePicker)
        {
            _nativeClip = new Panel
            {
                BackColor = AppTheme.Current.Surface,
                Margin = Padding.Empty,
                TabStop = false
            };
            _nativeClip.Controls.Add(input);
            Controls.Add(_nativeClip);
        }
        else
        {
            input.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(input);
        }
    }

    public override Size GetPreferredSize(Size proposedSize) =>
        new(proposedSize.Width > 0 ? proposedSize.Width : Width, AppTheme.ControlHeight);

    protected override void OnLayout(LayoutEventArgs levent)
    {
        base.OnLayout(levent);
        if (_nativeClip is not null)
        {
            var clipInset = BorderThickness + 5;
            var clipHeight = Math.Max(1, _input.PreferredSize.Height - 4);
            _nativeClip.Bounds = new Rectangle(
                clipInset,
                Math.Max(BorderThickness, (ClientSize.Height - clipHeight) / 2),
                Math.Max(0, ClientSize.Width - (clipInset * 2)),
                clipHeight);
            _input.Bounds = new Rectangle(
                -2,
                -2,
                _nativeClip.ClientSize.Width + 4,
                _input.PreferredSize.Height);
            return;
        }

        var horizontalInset = BorderThickness + 8;
        _input.Height = Math.Min(
            Math.Max(0, ClientSize.Height - (BorderThickness * 2)),
            _input.PreferredSize.Height);
        _input.Width = Math.Max(0, ClientSize.Width - (horizontalInset * 2));
        _input.Location = new Point(
            horizontalInset,
            Math.Max(BorderThickness, (ClientSize.Height - _input.PreferredSize.Height) / 2));
    }

    protected override void OnResize(EventArgs eventargs)
    {
        base.OnResize(eventargs);
        UpdateHostRegion();
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        base.OnMouseDown(e);
        _input.Focus();
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        var palette = AppTheme.Current;
        var inset = Math.Max(1, (int)Math.Ceiling(BorderThickness / 2F));
        var bounds = new Rectangle(
            inset,
            inset,
            Math.Max(0, Width - (inset * 2) - 1),
            Math.Max(0, Height - (inset * 2) - 1));

        if (SystemInformation.HighContrast)
        {
            using var highContrastPen = new Pen(SystemColors.ControlDark, BorderThickness);
            e.Graphics.DrawRectangle(highContrastPen, bounds);
            return;
        }

        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        using var path = RoundedRectangle(bounds, CornerRadius);
        using var background = new SolidBrush(palette.Surface);
        using var border = new Pen(
            !_input.Enabled ? palette.Border : _input.Focused ? palette.Accent : palette.Border,
            BorderThickness);
        e.Graphics.FillPath(background, path);
        e.Graphics.DrawPath(border, path);
    }

    private void UpdateHostRegion()
    {
        var oldRegion = Region;
        if (SystemInformation.HighContrast || Width <= 0 || Height <= 0)
        {
            Region = null;
        }
        else
        {
            using var path = RoundedRectangle(new Rectangle(0, 0, Width, Height), CornerRadius);
            Region = new Region(path);
        }
        oldRegion?.Dispose();
    }

    private static GraphicsPath RoundedRectangle(Rectangle bounds, int radius)
    {
        var path = new GraphicsPath();
        var diameter = Math.Min(Math.Max(0, radius * 2), Math.Min(bounds.Width, bounds.Height));
        if (diameter <= 0)
        {
            path.AddRectangle(bounds);
            return path;
        }

        var arc = new Rectangle(bounds.Location, new Size(diameter, diameter));
        path.AddArc(arc, 180, 90);
        arc.X = bounds.Right - diameter;
        path.AddArc(arc, 270, 90);
        arc.Y = bounds.Bottom - diameter;
        path.AddArc(arc, 0, 90);
        arc.X = bounds.Left;
        path.AddArc(arc, 90, 90);
        path.CloseFigure();
        return path;
    }
}
