namespace MorgenAttendanceSummaryCreator.Models;

public enum ProcessingStatus
{
    None,
    UploadStart,
    UploadEnd,
    CreateSummaryStart,
    CreateSummaryEnd
}