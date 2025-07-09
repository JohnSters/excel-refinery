using System.ComponentModel.DataAnnotations;

namespace ExcelRefinery.Models
{
    public class FileAnalysisResult
    {
        public string FileId { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public long FileSize { get; set; }
        public string FileType { get; set; } = string.Empty;
        public DateTime LastModified { get; set; }
        public string Status { get; set; } = "ready";
        public List<WorksheetInfo> Worksheets { get; set; } = new();
        public List<HeaderMapping> Headers { get; set; } = new();
        public List<string> ValidationErrors { get; set; } = new();
        public List<string> ValidationWarnings { get; set; } = new();
        public int QualityScore { get; set; }
    }

    public class WorksheetInfo
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public bool HasHeaders { get; set; }
        public bool Selected { get; set; }
        public List<string> DetectedHeaders { get; set; } = new();
        public string FirstDataRowPreview { get; set; } = string.Empty;
    }

    public class HeaderMapping
    {
        public string Id { get; set; } = string.Empty;
        public string DetectedName { get; set; } = string.Empty;
        public string StandardName { get; set; } = string.Empty;
        public string DataType { get; set; } = string.Empty;
        public bool Selected { get; set; }
        public bool IsRequired { get; set; }
        public double MatchConfidence { get; set; }
        public string Column { get; set; } = string.Empty;
        public string SampleData { get; set; } = string.Empty;
    }

    public class DataPreviewResult
    {
        public string FileId { get; set; } = string.Empty;
        public string WorksheetId { get; set; } = string.Empty;
        public List<string> Headers { get; set; } = new();
        public List<List<string>> Rows { get; set; } = new();
        public int TotalRows { get; set; }
        public bool HasMoreData { get; set; }
    }

    public class FileUploadRequest
    {
        [Required]
        public List<IFormFile> Files { get; set; } = new();
    }


} 