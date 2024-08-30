namespace MorgenAttendanceSummaryCreator.Models;

public class WorkingHoursData
{
    public decimal WorkingHours { get; set; } = 0;

    public decimal OvertimeHours { get; set; } = 0;

    public int LateInMinutes { get; set; } = 0;

    public int EarlyOutMinutes { get; set; } = 0;
}
