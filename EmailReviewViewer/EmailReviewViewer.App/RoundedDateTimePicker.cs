namespace EmailReviewViewer.App;

public sealed class RoundedDateTimePicker : DateTimePicker
{
    private const int WsBorder = 0x00800000;
    private const int WsExClientEdge = 0x00000200;

    protected override CreateParams CreateParams
    {
        get
        {
            var parameters = base.CreateParams;
            parameters.Style &= ~WsBorder;
            parameters.ExStyle &= ~WsExClientEdge;
            return parameters;
        }
    }
}
