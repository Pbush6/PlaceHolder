namespace EmailReviewViewer.App;

public sealed record AppPalette(
    Color Primary,
    Color Accent,
    Color Canvas,
    Color Surface,
    Color SubtleSurface,
    Color Border,
    Color Text,
    Color MutedText,
    Color Selection,
    Color SelectionText);

public static class AppTheme
{
    public const int ControlHeight = 32;
    public const int SectionPadding = 16;
    public const int SplitterWidth = 5;

    public static AppPalette Current => Resolve(SystemInformation.HighContrast);

    public static AppPalette Resolve(bool highContrast) =>
        highContrast
            ? new(
                SystemColors.Control,
                SystemColors.Highlight,
                SystemColors.Control,
                SystemColors.Window,
                SystemColors.Control,
                SystemColors.ControlDark,
                SystemColors.WindowText,
                SystemColors.GrayText,
                SystemColors.Highlight,
                SystemColors.HighlightText)
            : new(
                Color.FromArgb(27, 48, 74),
                Color.FromArgb(45, 105, 166),
                Color.FromArgb(241, 244, 248),
                Color.White,
                Color.FromArgb(247, 249, 252),
                Color.FromArgb(211, 218, 227),
                Color.FromArgb(31, 41, 55),
                Color.FromArgb(96, 108, 124),
                Color.FromArgb(218, 232, 247),
                Color.FromArgb(20, 45, 72));

    public static Font Font(float size = 9F, FontStyle style = FontStyle.Regular) =>
        new("Segoe UI", size, style, GraphicsUnit.Point);

    public static void StyleButton(Button button, ButtonKind kind)
    {
        var palette = Current;
        button.AutoSize = false;
        button.Height = ControlHeight;
        button.Padding = new Padding(12, 0, 12, 0);
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderSize = 1;
        button.Cursor = Cursors.Hand;
        button.UseVisualStyleBackColor = false;

        if (kind == ButtonKind.Primary)
        {
            button.BackColor = palette.Accent;
            button.ForeColor = SystemInformation.HighContrast ? SystemColors.HighlightText : Color.White;
            button.FlatAppearance.BorderColor = palette.Accent;
        }
        else if (kind == ButtonKind.Secondary)
        {
            button.BackColor = palette.Surface;
            button.ForeColor = palette.Text;
            button.FlatAppearance.BorderColor = palette.Border;
        }
        else
        {
            button.BackColor = palette.Surface;
            button.ForeColor = palette.MutedText;
            button.FlatAppearance.BorderSize = 0;
        }
    }

    public static void StyleInput(Control control)
    {
        control.Font = Font();
        control.BackColor = Current.Surface;
        control.ForeColor = Current.Text;
        control.MinimumSize = new Size(0, ControlHeight);
        control.Margin = new Padding(0, 3, 12, 3);
    }
}

public enum ButtonKind
{
    Primary,
    Secondary,
    Quiet
}
