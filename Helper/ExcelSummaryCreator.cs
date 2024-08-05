using MorgenAttendanceSummaryCreator.Models;
using Syncfusion.XlsIO;

namespace MorgenAttendanceSummaryCreator.Helper;

public class ExcelSummaryCreator
{
    public MemoryStream CreateSummary(string inputFileName)
    {
        using ExcelEngine excelEngine = new();
        IApplication application = excelEngine.Excel;

        application.DefaultVersion = ExcelVersion.Excel97to2003;

        FileStream inputStream = new(inputFileName, FileMode.Open);
        IWorkbook inputWorkbook = application.Workbooks.Open(inputStream);
        IWorksheet inputWorksheet = inputWorkbook.Worksheets[0];
        List<AttendanceSummaryData> summaryData = ReadAttendanceXls(inputWorksheet);
        StartAndEndDateInfo bounds = ParseStartAndEndDateInfo(inputWorksheet);
        IWorkbook outputWorkbook = application.Workbooks.Create(1);
        IWorksheet outputWorksheet = outputWorkbook.Worksheets[0];

        outputWorksheet.IsGridLinesVisible = false;
        WriteSummaryHeader(outputWorksheet, bounds);

        for (int index = 0; index < summaryData.Count; index++)
        {
            WriteSummaryDetail(outputWorksheet, summaryData[index], index);
        }

        outputWorksheet.UsedRange.AutofitColumns();
        outputWorksheet.UsedRange.AutofitRows();

        MemoryStream outputStream = new();

        outputWorkbook.SaveAs(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }

    private void WriteSummaryHeader(IWorksheet outputWorksheet, StartAndEndDateInfo bounds)
    {
        outputWorksheet.Range[1, 1].Text = "SUMMARY ABSENSI";
        outputWorksheet.Range[1, 1].CellStyle.Font.Size = 18;
        outputWorksheet.Range[1, 1].CellStyle.Font.Bold = true;
        outputWorksheet.Range[1, 1].RowHeight = 22;
        outputWorksheet.Range[1, 1].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
        outputWorksheet.Range[1, 1].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
        outputWorksheet.Range[1, 1, 1, 5].Merge();
        outputWorksheet.Range[2, 1].Text = "PERIODE: ";
        outputWorksheet.Range[2, 2].DateTime = bounds.StartDate;
        outputWorksheet.Range[2, 3].Text = " S/D ";
        outputWorksheet.Range[2, 4].DateTime = bounds.EndDate;
        outputWorksheet.Range[3, 1].Text = "Nomor";
        outputWorksheet.Range[3, 2].Text = "ID";
        outputWorksheet.Range[3, 3].Text = "Nama";
        outputWorksheet.Range[3, 4].Text = "Departemen";
        outputWorksheet.Range[3, 5].Text = "Jumlah Kehadiran";
    }

    private void WriteSummaryDetail(IWorksheet outputWorksheet, AttendanceSummaryData item, int index)
    {
        outputWorksheet.Range[index + 4, 1].Number = index + 1;
        outputWorksheet.Range[index + 4, 2].Text = item.EmployeeInfo.Id;
        outputWorksheet.Range[index + 4, 3].Text = item.EmployeeInfo.Name;
        outputWorksheet.Range[index + 4, 4].Text = item.EmployeeInfo.Department;
        outputWorksheet.Range[index + 4, 5].Number = item.AttendanceCount;
    }

    private List<AttendanceSummaryData> ReadAttendanceXls(IWorksheet inputWorksheet)
    {
        int rowCount = inputWorksheet.Rows.Length;
        int columnCount = inputWorksheet.Columns.Length;
        List<AttendanceSummaryData> summaryData = new();

        // only need to read once for day header 
        // string dayHeader = inputWorksheet.Range[3, 1].Text;

        for (int row = 2; row <= rowCount; row += 4)
        {
            summaryData.Add(
                new AttendanceSummaryData()
                {
                    EmployeeInfo = ParseEmployeeInfo(inputWorksheet, row),
                    AttendanceCount = CountEmployeeAttendance(inputWorksheet, row, columnCount)
                }
            );
        }

        return summaryData;
    }

    private StartAndEndDateInfo ParseStartAndEndDateInfo(IWorksheet inputWorksheet)
    {
        string periodInfo = inputWorksheet.Range[1, 1].Text.Split(' ', StringSplitOptions.TrimEntries)[0];
        string[] periodParts = periodInfo.Split([':', '/', '-']);
        int year = int.Parse(periodParts[0]);
        int startMonth = int.Parse(periodParts[1]);
        int startDay = int.Parse(periodParts[2]);
        int endMonth = int.Parse(periodParts[3]);
        int endDay = int.Parse(periodParts[4]);

        return new StartAndEndDateInfo()
        {
            StartDate = new DateTime(year, startMonth, startDay),
            EndDate = new DateTime(year, endMonth, endDay)
        };
    }
    private EmployeeInfo ParseEmployeeInfo(IWorksheet inputWorksheet, int row)
    {
        string[] employeeParts = inputWorksheet.Range[row, 1].Text.Split([' ', ':'], StringSplitOptions.TrimEntries);

        return new EmployeeInfo()
        {
            Id = employeeParts[1],
            Name = employeeParts[4],
            Department = employeeParts[7],
            Shift = employeeParts[13]
        };
    }
    private int CountEmployeeAttendance(IWorksheet inputWorksheet, int row, int columnCount)
    {
        int attendanceCount = 0;

        for (int column = 1; column <= columnCount; column++)
        {
            if (inputWorksheet.Range[row + 2, column].IsBlank ||
                string.IsNullOrEmpty(inputWorksheet.Range[row + 2, column].Text.Trim()))
            {
                continue;
            }

            attendanceCount++;
        }

        return attendanceCount;
    }

    private void WriteEmployeeTimeInAndOutHeader(IWorksheet inputWorksheet, IWorksheet outputWorksheet, StartAndEndDateInfo bounds)
    {
        outputWorksheet.Range[1, 1].Text = "DATA WAKTU MASUK/KELUAR KARYAWAN";
        outputWorksheet.Range[1, 1].CellStyle.Font.Size = 18;
        outputWorksheet.Range[1, 1].CellStyle.Font.Bold = true;
        outputWorksheet.Range[1, 1].RowHeight = 22;
        outputWorksheet.Range[1, 1].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
        outputWorksheet.Range[1, 1].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
        outputWorksheet.Range[1, 1, 1, 2 * inputWorksheet.Columns.Length + 4].Merge();
        outputWorksheet.Range[2, 1].Text = "PERIODE: ";
        outputWorksheet.Range[2, 2].DateTime = bounds.StartDate;
        outputWorksheet.Range[2, 3].Text = " S/D ";
        outputWorksheet.Range[2, 4].DateTime = bounds.EndDate;
        outputWorksheet.Range[3, 1].Text = "Nomor";
        outputWorksheet.Range[3, 2].Text = "ID";
        outputWorksheet.Range[3, 3].Text = "Nama";
        outputWorksheet.Range[3, 4].Text = "Departemen";

        outputWorksheet.Range[3, 1, 4, 1].Merge();
        outputWorksheet.Range[3, 2, 4, 2].Merge();
        outputWorksheet.Range[3, 3, 4, 3].Merge();
        outputWorksheet.Range[3, 4, 4, 4].Merge();

        for (int column = 1; column <= inputWorksheet.Columns.Length; column++)
        {
            outputWorksheet.Range[3, 3 + column * 2].DateTime = bounds.StartDate.AddDays(column - 1);
            outputWorksheet.Range[3, 3 + column * 2, 3, 4 + column * 2].Merge();
            outputWorksheet.Range[4, 3 + column * 2].Text = "IN";
            outputWorksheet.Range[4, 4 + column * 2].Text = "OUT";
        }
    }
    private void WriteEmployeeTimeInAndOutDetail(IWorksheet outputWorksheet, EmployeeTimeInAndOutData item, int index)
    {
        outputWorksheet.Range[index + 5, 1].Number = index + 1;
        outputWorksheet.Range[index + 5, 2].Text = item.EmployeeInfo.Id;
        outputWorksheet.Range[index + 5, 3].Text = item.EmployeeInfo.Name;
        outputWorksheet.Range[index + 5, 4].Text = item.EmployeeInfo.Department;
        DateOnly dummyDate = new DateOnly(2000, 1, 1);

        for (int column = 0; column < item.TimeInAndOutData.Count; column++)
        {
            outputWorksheet.Range[index + 5, 5 + column * 2].DateTime = dummyDate.ToDateTime(item.TimeInAndOutData[column].TimeIn);
            outputWorksheet.Range[index + 5, 6 + column * 2].DateTime = dummyDate.ToDateTime(item.TimeInAndOutData[column].TimeOut);
        }
    }

    private List<EmployeeTimeInAndOutData> ParseEmployeeTimeInAndOutDataFromXls(IWorksheet inputWorksheet)
    {
        int rowCount = inputWorksheet.Rows.Length;
        int columnCount = inputWorksheet.Columns.Length;
        List<EmployeeTimeInAndOutData> list = [];

        for (int row = 2; row <= rowCount; row += 4)
        {
            list.Add(
                new EmployeeTimeInAndOutData
                {
                    EmployeeInfo = ParseEmployeeInfo(inputWorksheet, row),
                    TimeInAndOutData = RetrieveTimeInAndOutList(inputWorksheet, row, columnCount)
                }
            );
        }

        return list;
    }
    private List<TimeInAndOutData> RetrieveTimeInAndOutList(IWorksheet inputWorksheet, int row, int columnCount)
    {
        List<TimeInAndOutData> result = [];

        for (int column = 1; column <= columnCount; column++)
        {
            result.Add(ParseTimeInAndOut(inputWorksheet.Range[row + 2, column]));
        }

        return result;
    }
    private TimeInAndOutData ParseTimeInAndOut(IRange range)
    {
        if (range.IsBlank || string.IsNullOrWhiteSpace(range.Text))
        {
            return new TimeInAndOutData();
        }

        string[] clockParts = range.Text.Split([' '], StringSplitOptions.TrimEntries);
        TimeInAndOutData result = new();

        if (!string.IsNullOrWhiteSpace(clockParts[0]))
        {
            result.TimeIn = TimeOnly.Parse(clockParts[0]);
        }

        if (!string.IsNullOrWhiteSpace(clockParts[1]))
        {
            result.TimeOut = TimeOnly.Parse(clockParts[1]);
        }

        return result;
    }
}