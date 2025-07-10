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
        public DuplicateFileInfo? DuplicateInfo { get; set; }
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

    public class DuplicateFileInfo
    {
        public bool IsDuplicate { get; set; }
        public List<string> DuplicateFileNames { get; set; } = new();
        public List<string> DuplicateFileIds { get; set; } = new();
        public string WorksheetName { get; set; } = string.Empty;
        public double SimilarityScore { get; set; }
        public DateTime DetectedAt { get; set; } = DateTime.Now;
        public string ComparisonDetails { get; set; } = string.Empty;
        public SimilarityLevel SimilarityLevel { get; set; } = SimilarityLevel.Different;
        public ComparisonResult DetailedComparison { get; set; } = new();
    }

    public enum SimilarityLevel
    {
        ExactMatch = 100,        // 100% match - Perfect match
        NearIdentical = 99,      // 98% - 99.9% match - Near perfect  
        HighSimilarity = 95,     // 90% - 97.9% match - Good consistency
        ModerateSimilarity = 80, // 80% - 89.9% match - Some differences
        LowSimilarity = 50,      // 50% - 79.9% match - Significant differences
        Different = 0            // < 50% match - Very different
    }

    /// <summary>
    /// Represents the status type for comparison results
    /// </summary>
    public enum ComparisonStatus
    {
        Success,    // 90-100% similarity - Good data consistency
        Warning,    // 50-89% similarity - Some differences found
        Error       // <50% similarity - Significant differences
    }

    public class ComparisonResult
    {
        public int TotalRows { get; set; }
        public int MatchingRows { get; set; }
        public int DifferentRows { get; set; }
        public List<string> DifferentRowSamples { get; set; } = new();
        public bool HeadersMatch { get; set; } = true;
        public List<string> MissingHeaders { get; set; } = new();
        public List<string> ExtraHeaders { get; set; } = new();
        public List<string> MatchedHeaders { get; set; } = new();
        public List<HeaderMatchInfo> HeaderMatches { get; set; } = new();
        public string SummaryMessage { get; set; } = string.Empty;
        public ComparisonStatus Status { get; set; } = ComparisonStatus.Error;
        public double SimilarityPercentage { get; set; }
        public SimilarityLevel SimilarityLevel { get; set; } = SimilarityLevel.Different;
    }

    /// <summary>
    /// Detailed information about how headers were matched between worksheets
    /// </summary>
    public class HeaderMatchInfo
    {
        public string SourceHeader { get; set; } = string.Empty;
        public string TargetHeader { get; set; } = string.Empty;
        public double MatchConfidence { get; set; }
        public string MatchReason { get; set; } = string.Empty; // "exact", "normalized", "fuzzy"
    }

    /// <summary>
    /// Represents a worksheet comparison request
    /// </summary>
    public class WorksheetComparisonRequest
    {
        public string File1Id { get; set; } = string.Empty;
        public string File1WorksheetName { get; set; } = string.Empty;
        public string File2Id { get; set; } = string.Empty;
        public string File2WorksheetName { get; set; } = string.Empty;
        public double MatchThreshold { get; set; } = 0.90; // 90% threshold for row matching
    }

    /// <summary>
    /// Enhanced file integrity result with worksheet-specific comparisons
    /// </summary>
    public class FileIntegrityResult
    {
        public string FileId { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public List<WorksheetIntegrityComparison> WorksheetComparisons { get; set; } = new();
        public bool HasExactMatches { get; set; }
        public bool HasGoodMatches { get; set; } // 90%+ matches
        public bool HasIssues { get; set; }
        public string OverallStatus { get; set; } = string.Empty; // "excellent_match", "good_match", "has_differences", "poor_match"
        public ComparisonStatus Status { get; set; } = ComparisonStatus.Error;
    }

    /// <summary>
    /// Worksheet-level comparison details
    /// </summary>
    public class WorksheetIntegrityComparison
    {
        public string SourceWorksheetName { get; set; } = string.Empty;
        public string ComparedWithFileId { get; set; } = string.Empty;
        public string ComparedWithFileName { get; set; } = string.Empty;
        public string ComparedWithWorksheetName { get; set; } = string.Empty;
        public double SimilarityScore { get; set; }
        public SimilarityLevel SimilarityLevel { get; set; } = SimilarityLevel.Different;
        public ComparisonStatus Status { get; set; } = ComparisonStatus.Error;
        public ComparisonResult DetailedComparison { get; set; } = new();
        public List<string> SpecificDifferences { get; set; } = new();
        public string StatusMessage { get; set; } = string.Empty;
        public string StatusIcon { get; set; } = string.Empty; // For UI display
        public string StatusColor { get; set; } = string.Empty; // For UI display
    }

    // Keep existing for backward compatibility
    public class FileIntegrityComparison
    {
        public string ComparedWithFileId { get; set; } = string.Empty;
        public string ComparedWithFileName { get; set; } = string.Empty;
        public double SimilarityScore { get; set; }
        public bool IsExactMatch { get; set; }
        public ComparisonResult DetailedComparison { get; set; } = new();
        public List<string> SpecificDifferences { get; set; } = new();
        public string WorksheetName { get; set; } = string.Empty;
        public string StatusMessage { get; set; } = string.Empty;
    }

    public class ProcessedFileCache
    {
        public string FileId { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public DateTime ProcessedAt { get; set; } = DateTime.Now;
        public List<NormalizedWorksheetData> Worksheets { get; set; } = new();
        public string FileHash { get; set; } = string.Empty;
    }

    public class NormalizedWorksheetData
    {
        public string Name { get; set; } = string.Empty;
        public List<string> Headers { get; set; } = new();
        public List<Dictionary<string, string>> Rows { get; set; } = new();
        public string DataHash { get; set; } = string.Empty;
        public int OriginalRowCount { get; set; } // Before filtering
        public int DataRowCount { get; set; } // After filtering out header/filter rows
    }

    /// <summary>
    /// Represents a data row for comparison with position-independent matching
    /// </summary>
    public class DataRowComparison
    {
        public Dictionary<string, string> Data { get; set; } = new();
        public int OriginalRowNumber { get; set; }
        public string RowHash { get; set; } = string.Empty; // For quick comparison
        public List<string> KeyFields { get; set; } = new(); // Primary identifying fields
    }
} 