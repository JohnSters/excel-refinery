/**
 * ExcelProcessingService.cs
 * Main orchestration service for Excel file processing and integrity checking
 * Uses specialized services for analysis, normalization, and comparison
 * Author: ExcelRefinery System
 */

using ExcelRefinery.Models;
using System.Collections.Concurrent;

namespace ExcelRefinery.Services
{
    public interface IExcelProcessingService
    {
        Task<FileAnalysisResult> AnalyzeFileAsync(IFormFile file);
        Task<DataPreviewResult> GetDataPreviewAsync(string filePath, string worksheetName, int maxRows = 10);
        Task<List<HeaderMapping>> DetectAndMapHeadersAsync(string filePath, string worksheetName);
        Task<List<FileIntegrityResult>> CheckFileIntegrityAsync(List<string> fileIds);
        Task<List<WorksheetIntegrityComparison>> CheckWorksheetIntegrityAsync(List<WorksheetComparisonRequest> requests);
        void CleanupOldTempFiles(int maxAgeHours = 24);
        void ClearProcessedFileCache();
    }

    public class ExcelProcessingService : IExcelProcessingService
    {
        private readonly ILogger<ExcelProcessingService> _logger;
        private readonly IFileAnalysisService _fileAnalysisService;
        private readonly IDataNormalizationService _dataNormalizationService;
        private readonly IWorksheetComparisonService _worksheetComparisonService;
        
        // In-memory cache for processed files to detect duplicates and enable comparisons
        private static readonly ConcurrentDictionary<string, ProcessedFileCache> _processedFilesCache = new();
        private static readonly object _cacheLock = new object();

        public ExcelProcessingService(
            ILogger<ExcelProcessingService> logger,
            IFileAnalysisService fileAnalysisService,
            IDataNormalizationService dataNormalizationService,
            IWorksheetComparisonService worksheetComparisonService)
        {
            _logger = logger;
            _fileAnalysisService = fileAnalysisService;
            _dataNormalizationService = dataNormalizationService;
            _worksheetComparisonService = worksheetComparisonService;
        }

        public async Task<FileAnalysisResult> AnalyzeFileAsync(IFormFile file)
        {
            try
            {
                _logger.LogInformation("=== Starting File Analysis Orchestration ===");
                _logger.LogInformation("Processing file: {FileName} ({FileSize} bytes)", file.FileName, file.Length);

                // Step 1: Analyze the file structure and content
                var analysisResult = await _fileAnalysisService.AnalyzeFileAsync(file);
                
                if (analysisResult.Status == "error")
                {
                    _logger.LogError("File analysis failed for {FileName}: {Errors}", 
                        file.FileName, string.Join(", ", analysisResult.ValidationErrors));
                    return analysisResult;
                }

                // Step 2: Normalize the data for comparison purposes (only if successful)
                var tempFileName = $"{analysisResult.FileId}_{Path.GetFileName(file.FileName)}";
                await CacheNormalizedFileDataAsync(tempFileName, analysisResult);

                _logger.LogInformation("File analysis orchestration complete for {FileName}: {WorksheetCount} worksheets processed", 
                    file.FileName, analysisResult.Worksheets.Count);

                return analysisResult;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in file analysis orchestration for {FileName}", file.FileName);
                
                return new FileAnalysisResult
                {
                    FileId = Guid.NewGuid().ToString(),
                    FileName = file.FileName,
                    FileSize = file.Length,
                    FileType = file.ContentType,
                    LastModified = DateTime.Now,
                    Status = "error",
                    ValidationErrors = new List<string> { $"Orchestration error: {ex.Message}" }
                };
            }
        }

        public async Task<DataPreviewResult> GetDataPreviewAsync(string filePath, string worksheetName, int maxRows = 10)
        {
            return await _fileAnalysisService.GetDataPreviewAsync(filePath, worksheetName, maxRows);
        }

        public async Task<List<HeaderMapping>> DetectAndMapHeadersAsync(string filePath, string worksheetName)
        {
            return await _fileAnalysisService.DetectAndMapHeadersAsync(filePath, worksheetName);
        }

        public async Task<List<WorksheetIntegrityComparison>> CheckWorksheetIntegrityAsync(List<WorksheetComparisonRequest> requests)
        {
            _logger.LogInformation("=== Starting Worksheet-Specific Integrity Check ===");
            _logger.LogInformation("Processing {RequestCount} worksheet comparison requests", requests.Count);

            try
            {
                List<ProcessedFileCache> cachedFiles;
                
                lock (_cacheLock)
                {
                    var requestedFileIds = requests.SelectMany(r => new[] { r.File1Id, r.File2Id }).Distinct().ToList();
                    cachedFiles = _processedFilesCache.Values
                        .Where(f => requestedFileIds.Contains(f.FileId))
                        .ToList();
                        
                    _logger.LogInformation("Found {CachedCount} cached files for {RequestedCount} unique file IDs", 
                        cachedFiles.Count, requestedFileIds.Count);
                }

                if (cachedFiles.Count < 2)
                {
                    _logger.LogWarning("Insufficient cached files for worksheet integrity check: {Count}", cachedFiles.Count);
                    return new List<WorksheetIntegrityComparison>();
                }

                // Use the specialized comparison service for worksheet-level comparisons
                var comparisons = await _worksheetComparisonService.CompareWorksheetsBetweenFilesAsync(cachedFiles, requests);

                _logger.LogInformation("Worksheet integrity check complete: {ComparisonCount} comparisons performed", 
                    comparisons.Count);

                return comparisons;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error during worksheet integrity check");
                return new List<WorksheetIntegrityComparison>();
            }
        }

        public async Task<List<FileIntegrityResult>> CheckFileIntegrityAsync(List<string> fileIds)
        {
            _logger.LogInformation("=== Starting File-Level Integrity Check ===");
            _logger.LogInformation("Checking integrity for file IDs: [{FileIds}]", string.Join(", ", fileIds));

            var integrityResults = new List<FileIntegrityResult>();
            
            try
            {
                if (fileIds.Count < 2)
                {
                    _logger.LogWarning("File integrity check requires at least 2 files, received: {Count}", fileIds.Count);
                    return integrityResults;
                }

                List<ProcessedFileCache> cachedFiles;
                
                lock (_cacheLock)
                {
                    cachedFiles = _processedFilesCache.Values
                        .Where(f => fileIds.Contains(f.FileId))
                        .ToList();
                        
                    _logger.LogInformation("Found {FoundCount} cached files out of {RequestedCount} requested", 
                        cachedFiles.Count, fileIds.Count);
                }

                if (cachedFiles.Count < 2)
                {
                    _logger.LogWarning("Insufficient cached files for integrity check: {Count}", cachedFiles.Count);
                    return integrityResults;
                }

                // Generate comparison requests for all worksheet combinations
                var comparisonRequests = GenerateWorksheetComparisonRequests(cachedFiles);
                
                _logger.LogInformation("Generated {RequestCount} worksheet comparison requests", comparisonRequests.Count);

                if (!comparisonRequests.Any())
                {
                    _logger.LogWarning("No worksheet comparison requests generated");
                    return integrityResults;
                }

                // Perform worksheet-level comparisons
                var worksheetComparisons = await _worksheetComparisonService.CompareWorksheetsBetweenFilesAsync(cachedFiles, comparisonRequests);

                // Group results by source file for file-level integrity results
                var groupedComparisons = worksheetComparisons
                    .GroupBy(c => cachedFiles.First(f => f.Worksheets.Any(w => w.Name == c.SourceWorksheetName)).FileId)
                    .ToList();

                foreach (var fileGroup in groupedComparisons)
                {
                    var sourceFile = cachedFiles.First(f => f.FileId == fileGroup.Key);
                    
                    var integrityResult = new FileIntegrityResult
                    {
                        FileId = sourceFile.FileId,
                        FileName = sourceFile.FileName,
                        WorksheetComparisons = fileGroup.ToList()
                    };

                    // Determine overall file status based on worksheet comparisons
                    SetFileIntegrityStatus(integrityResult);

                    integrityResults.Add(integrityResult);
                    
                    _logger.LogInformation("File '{FileName}' integrity analysis: Status={Status}, Comparisons={Count}", 
                        sourceFile.FileName, integrityResult.OverallStatus, integrityResult.WorksheetComparisons.Count);
                }

                _logger.LogInformation("File-level integrity check complete: {ResultCount} files analyzed", 
                    integrityResults.Count);

                return integrityResults;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error during file integrity check");
                return integrityResults;
            }
        }

        private List<WorksheetComparisonRequest> GenerateWorksheetComparisonRequests(List<ProcessedFileCache> files)
        {
            var requests = new List<WorksheetComparisonRequest>();
            
            try
            {
                // Compare each file with every other file
                for (int i = 0; i < files.Count; i++)
                {
                    for (int j = i + 1; j < files.Count; j++)
                    {
                        var file1 = files[i];
                        var file2 = files[j];
                        
                        // Try to match worksheets by name similarity
                        foreach (var worksheet1 in file1.Worksheets)
                        {
                            // Look for similar worksheet names in file2
                            var matchingWorksheet = file2.Worksheets
                                .OrderByDescending(w => CalculateNameSimilarity(worksheet1.Name, w.Name))
                                .FirstOrDefault();
                            
                            if (matchingWorksheet != null)
                            {
                                requests.Add(new WorksheetComparisonRequest
                                {
                                    File1Id = file1.FileId,
                                    File1WorksheetName = worksheet1.Name,
                                    File2Id = file2.FileId,
                                    File2WorksheetName = matchingWorksheet.Name,
                                    MatchThreshold = 0.90 // 90% threshold for row matching
                                });
                                
                                _logger.LogDebug("Generated comparison request: {File1}[{WS1}] vs {File2}[{WS2}]", 
                                    file1.FileName, worksheet1.Name, file2.FileName, matchingWorksheet.Name);
                            }
                        }
                    }
                }
                
                _logger.LogInformation("Generated {RequestCount} worksheet comparison requests from {FileCount} files", 
                    requests.Count, files.Count);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating worksheet comparison requests");
            }
            
            return requests;
        }

        private double CalculateNameSimilarity(string name1, string name2)
        {
            if (string.Equals(name1, name2, StringComparison.OrdinalIgnoreCase))
                return 1.0;
                
            // Simple similarity based on common words and length
            var words1 = name1.ToLowerInvariant().Split('_', ' ', '-').Where(w => !string.IsNullOrEmpty(w)).ToHashSet();
            var words2 = name2.ToLowerInvariant().Split('_', ' ', '-').Where(w => !string.IsNullOrEmpty(w)).ToHashSet();
            
            if (!words1.Any() || !words2.Any())
                return 0.0;
            
            var intersection = words1.Intersect(words2).Count();
            var union = words1.Union(words2).Count();
            
            return union > 0 ? (double)intersection / union : 0.0;
        }

        private void SetFileIntegrityStatus(FileIntegrityResult result)
        {
            if (!result.WorksheetComparisons.Any())
            {
                result.OverallStatus = "no_comparison";
                result.Status = ComparisonStatus.Error;
                return;
            }

            var successCount = result.WorksheetComparisons.Count(c => c.Status == ComparisonStatus.Success);
            var warningCount = result.WorksheetComparisons.Count(c => c.Status == ComparisonStatus.Warning);
            var errorCount = result.WorksheetComparisons.Count(c => c.Status == ComparisonStatus.Error);
            
            result.HasExactMatches = result.WorksheetComparisons.Any(c => c.SimilarityLevel == SimilarityLevel.ExactMatch);
            result.HasGoodMatches = result.WorksheetComparisons.Any(c => c.Status == ComparisonStatus.Success);
            result.HasIssues = warningCount > 0 || errorCount > 0;

            // Determine overall status based on worksheet comparison results
            if (successCount == result.WorksheetComparisons.Count)
            {
                // All comparisons are successful (90%+ similarity)
                if (result.HasExactMatches)
                {
                    result.OverallStatus = "excellent_match";
                    result.Status = ComparisonStatus.Success;
                }
                else
                {
                    result.OverallStatus = "good_match";
                    result.Status = ComparisonStatus.Success;
                }
            }
            else if (result.HasGoodMatches)
            {
                // Some good matches, some issues
                result.OverallStatus = "has_differences";
                result.Status = ComparisonStatus.Warning;
            }
            else
            {
                // No good matches
                result.OverallStatus = "poor_match";
                result.Status = ComparisonStatus.Error;
            }
            
            _logger.LogDebug("File status determination: Success={Success}, Warning={Warning}, Error={Error} -> Status={Status}", 
                successCount, warningCount, errorCount, result.OverallStatus);
        }

        private async Task CacheNormalizedFileDataAsync(string filePath, FileAnalysisResult analysisResult)
        {
            try
            {
                _logger.LogInformation("Caching normalized data for file: {FileName}", analysisResult.FileName);

                var normalizedData = await _dataNormalizationService.NormalizeFileDataAsync(filePath, analysisResult);
                var fileHash = await _dataNormalizationService.CalculateFileHashAsync(filePath);
                
                var cacheEntry = new ProcessedFileCache
                {
                    FileId = analysisResult.FileId,
                    FileName = analysisResult.FileName,
                    ProcessedAt = DateTime.Now,
                    Worksheets = normalizedData,
                    FileHash = fileHash
                };
                
                lock (_cacheLock)
                {
                    _processedFilesCache[analysisResult.FileId] = cacheEntry;
                    
                    // Clean old entries (keep only last 100 files to prevent memory issues)
                    if (_processedFilesCache.Count > 100)
                    {
                        var oldestEntries = _processedFilesCache.Values
                            .OrderBy(v => v.ProcessedAt)
                            .Take(_processedFilesCache.Count - 100)
                            .Select(v => v.FileId)
                            .ToList();
                            
                        foreach (var key in oldestEntries)
                        {
                            _processedFilesCache.TryRemove(key, out _);
                        }
                        
                        _logger.LogInformation("Cleaned {CleanedCount} old cache entries", oldestEntries.Count);
                    }
                }
                
                _logger.LogInformation("Successfully cached file data for {FileName}: {WorksheetCount} worksheets, Hash={Hash}", 
                    analysisResult.FileName, normalizedData.Count, fileHash[..8] + "...");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error caching normalized file data for {FileName}", analysisResult.FileName);
            }
        }

        public void CleanupOldTempFiles(int maxAgeHours = 24)
        {
            _fileAnalysisService.CleanupOldTempFiles(maxAgeHours);
        }

        public void ClearProcessedFileCache()
        {
            lock (_cacheLock)
            {
                var cachedCount = _processedFilesCache.Count;
                _processedFilesCache.Clear();
                _logger.LogInformation("Cleared processed file cache: {CachedCount} entries removed", cachedCount);
            }
        }
    }
} 