namespace MorgenAttendanceSummaryCreator.Models;

public class DailyAttendanceData
{
    public DateTime ClockIn { get; set; } = DateTime.MinValue;
    public DateTime ClockOut { get; set; } = DateTime.MinValue;
}
