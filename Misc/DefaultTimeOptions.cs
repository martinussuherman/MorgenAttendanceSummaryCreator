namespace MorgenAttendanceSummaryCreator.Misc;

public class DefaultTimeOptions
{
    public const string OptionsName = "DefaultTime";

    public TimeOnly In { get; set; } = TimeOnly.MinValue;

    public TimeOnly Out { get; set; } = TimeOnly.MinValue;
}
