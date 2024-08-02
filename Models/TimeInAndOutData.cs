namespace MorgenAttendanceSummaryCreator.Models;

public class TimeInAndOutData
{
    public TimeOnly TimeIn { get; set; } = TimeOnly.MinValue;

    public TimeOnly TimeOut { get; set; } = TimeOnly.MinValue;
}
