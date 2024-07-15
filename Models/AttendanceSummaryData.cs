namespace MorgenAttendanceSummaryCreator.Models;

public class AttendanceSummaryData
{
    public EmployeeInfo EmployeeInfo { get; set; } = new EmployeeInfo();

    public int AttendanceCount { get; set; } = 0;
}