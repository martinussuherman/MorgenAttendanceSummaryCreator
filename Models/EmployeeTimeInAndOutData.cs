namespace MorgenAttendanceSummaryCreator.Models;

public class EmployeeTimeInAndOutData
{
    public EmployeeInfo EmployeeInfo { get; set; } = new EmployeeInfo();

    public List<TimeInAndOutData> TimeInAndOutData { get; set; } = [];
}