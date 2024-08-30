namespace MorgenAttendanceSummaryCreator.Models;

public class EmployeeWorkingHoursData
{
    public EmployeeInfo EmployeeInfo { get; set; } = new EmployeeInfo();

    public List<WorkingHoursData> WorkingHoursData { get; set; } = [];
}