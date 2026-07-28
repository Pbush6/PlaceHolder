using System.Drawing.Drawing2D;

namespace EmailReviewViewer.App;

public sealed class RoundedButton : Button
{
    private bool _hovered;
    private bool _pressed;
    private int _cornerRadius = 7;
    private int _borderThickness = 2;

    public int CornerRadius
    {
        get => _cornerRadius;
        set
        {
            _cornerRadius = Math.Max(0, value);
            Invalidate();
        }
    }

    public int BorderThickness
    {
        get => _borderThickness;
        set
        {
            _borderThickness = Math.Max(0, value);
            Invalidate();
        }
    }

    public Color BorderColor { get; set; } = SystemColors.ControlDark;
    public Color HoverBackColor { get; set; } = SystemColors.ControlLight;
    public Color PressedBackColor { get; set; } = SystemColors.ControlDark;

    public RoundedButton()
    {
        SetStyle(
            ControlStyles.UserPaint |
            ControlStyles.AllPaintingInWmPaint |
            ControlStyles.OptimizedDoubleBuffer |
            ControlStyles.ResizeRedraw,
            true);
        FlatStyle = FlatStyle.Flat;
        FlatAppearance.BorderSize = 0;
        UseVisualStyleBackColor = false;
    }

    protected override void OnMouseEnter(EventArgs e)
    {
        base.OnMouseEnter(e);
        _hovered = true;
        Invalidate();
    }

    protected override void OnMouseLeave(EventArgs e)
    {
        base.OnMouseLeave(e);
        _hovered = false;
        _pressed = false;
        Invalidate();
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        base.OnMouseDown(e);
        if (e.Button == MouseButtons.Left)
        {
            _pressed = true;
            Invalidate();
        }
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        base.OnMouseUp(e);
        _pressed = false;
        Invalidate();
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        base.OnKeyDown(e);
        if (e.KeyCode is Keys.Space or Keys.Enter)
        {
            _pressed = true;
            Invalidate();
        }
    }

    protected override void OnKeyUp(KeyEventArgs e)
    {
        base.OnKeyUp(e);
        _pressed = false;
        Invalidate();
    }

    protected override void OnEnabledChanged(EventArgs e)
    {
        base.OnEnabledChanged(e);
        Invalidate();
    }

    protected override void OnResize(EventArgs e)
    {
        base.OnResize(e);
        Invalidate();
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        if (SystemInformation.HighContrast)
        {
            base.OnPaint(e);
            return;
        }

        e.Graphics.Clear(Parent?.BackColor ?? AppTheme.Current.Surface);
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        e.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
        var inset = Math.Max(0, (int)Math.Ceiling(BorderThickness / 2F));
        var bounds = new Rectangle(
            inset,
            inset,
            Math.Max(0, Width - (inset * 2) - 1),
            Math.Max(0, Height - (inset * 2) - 1));
        using var path = RoundedRectangle(bounds, CornerRadius);
        var background = !Enabled
            ? AppTheme.Current.SubtleSurface
            : _pressed
                ? PressedBackColor
                : _hovered
                    ? HoverBackColor
                    : BackColor;
        using var backgroundBrush = new SolidBrush(background);
        using var borderPen = new Pen(BorderColor, BorderThickness);
        e.Graphics.FillPath(backgroundBrush, path);
        if (BorderThickness > 0)
            e.Graphics.DrawPath(borderPen, path);

        var textColor = Enabled ? ForeColor : SystemColors.GrayText;
        TextRenderer.DrawText(
            e.Graphics,
            Text,
            Font,
            ClientRectangle,
            textColor,
            TextFormatFlags.HorizontalCenter |
            TextFormatFlags.VerticalCenter |
            TextFormatFlags.EndEllipsis |
            TextFormatFlags.SingleLine);

        if (Focused && ShowFocusCues)
            ControlPaint.DrawFocusRectangle(e.Graphics, Rectangle.Inflate(ClientRectangle, -4, -4), textColor, background);
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
