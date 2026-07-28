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
    public const int ControlHeight = 36;
    public const int SectionPadding = 16;
    public const int SplitterWidth = 5;
    public const int CornerRadius = 8;
    public const int BorderThickness = 2;
    private static readonly Lazy<string> BodyFontFamily = new(
        () => ResolveFontFamily("Segoe UI Variable Text", "Segoe UI"));
    private static readonly Lazy<string> DisplayFontFamily = new(
        () => ResolveFontFamily("Segoe UI Variable Display", "Segoe UI Semibold", "Segoe UI"));

    public static AppPalette Current => Resolve(SystemInformation.HighContrast);
    public static string BodyFontFamilyName => BodyFontFamily.Value;
    public static string DisplayFontFamilyName => DisplayFontFamily.Value;

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

    public static Font Font(float size = 9.5F, FontStyle style = FontStyle.Regular) =>
        new(BodyFontFamilyName, size, style, GraphicsUnit.Point);

    public static Font DisplayFont(float size = 10F, FontStyle style = FontStyle.Regular) =>
        new(DisplayFontFamilyName, size, style, GraphicsUnit.Point);

    public static void StyleButton(Button button, ButtonKind kind)
    {
        var palette = Current;
        button.AutoSize = false;
        button.Height = ControlHeight;
        button.Padding = new Padding(12, 0, 12, 0);
        button.Font = DisplayFont(10F, FontStyle.Bold);
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderSize = BorderThickness;
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
            button.ForeColor = palette.Text;
            button.FlatAppearance.BorderColor = palette.Border;
        }

        if (button is RoundedButton rounded)
        {
            rounded.CornerRadius = CornerRadius;
            rounded.BorderThickness = BorderThickness;
            rounded.BorderColor = button.FlatAppearance.BorderColor;
            rounded.HoverBackColor = kind switch
            {
                ButtonKind.Primary => Color.FromArgb(37, 91, 145),
                ButtonKind.Secondary => palette.SubtleSurface,
                _ => palette.SubtleSurface
            };
            rounded.PressedBackColor = kind switch
            {
                ButtonKind.Primary => Color.FromArgb(29, 74, 119),
                _ => palette.Selection
            };
        }
    }

    public static void StyleInput(Control control)
    {
        control.Font = Font();
        control.BackColor = Current.Surface;
        control.ForeColor = Current.Text;
        control.MinimumSize = new Size(0, ControlHeight);
        control.Margin = new Padding(0, 3, 12, 3);
        if (control is TextBoxBase textBox)
            textBox.BorderStyle = BorderStyle.None;
    }

    private static string ResolveFontFamily(params string[] candidates)
    {
        var installed = FontFamily.Families
            .Select(family => family.Name)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        return candidates.First(installed.Contains);
    }
}

public enum ButtonKind
{
    Primary,
    Secondary,
    Quiet
}
