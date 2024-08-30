namespace MorgenAttendanceSummaryCreator.Models;

public class EmployeeWorkingHoursData
{
    public EmployeeInfo EmployeeInfo { get; set; } = new EmployeeInfo();

    public List<WorkingHoursData> WorkingHoursData { get; set; } = [];

    public int NoTimeInCount { get; set; } = 0;

    public int NoTimeOutCount { get; set; } = 0;
}