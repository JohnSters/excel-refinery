using ClosedXML.Excel;
using ExcelRefinery.Models;
using System.Text.RegularExpressions;

namespace ExcelRefinery.Services
{
    public interface IExcelProcessingService
    {
        Task<FileAnalysisResult> AnalyzeFileAsync(IFormFile file);
        Task<DataPreviewResult> GetDataPreviewAsync(string filePath, string worksheetName, int maxRows = 10);
        Task<List<HeaderMapping>> DetectAndMapHeadersAsync(string filePath, string worksheetName);
        void CleanupOldTempFiles(int maxAgeHours = 24);
    }

    public class ExcelProcessingService : IExcelProcessingService
    {
        private readonly ILogger<ExcelProcessingService> _logger;
        private readonly string _tempFilePath;

        public ExcelProcessingService(ILogger<ExcelProcessingService> logger, IWebHostEnvironment environment)
        {
            _logger = logger;
            _tempFilePath = Path.Combine(environment.WebRootPath, "temp");
            
            // Ensure temp directory exists
            if (!Directory.Exists(_tempFilePath))
                Directory.CreateDirectory(_tempFilePath);
        }

        public async Task<FileAnalysisResult> AnalyzeFileAsync(IFormFile file)
        {
            var fileId = Guid.NewGuid().ToString();
            // Sanitize filename to prevent path traversal
            var sanitizedFileName = Path.GetFileName(file.FileName);
            var tempFileName = $"{fileId}_{sanitizedFileName}";
            var tempFilePath = Path.Combine(_tempFilePath, tempFileName);

            try
            {
                // Save uploaded file temporarily
                using (var stream = new FileStream(tempFilePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                var result = new FileAnalysisResult
                {
                    FileId = fileId,
                    FileName = file.FileName,
                    FileSize = file.Length,
                    FileType = file.ContentType,
                    LastModified = DateTime.Now
                };

                // Determine file type and process accordingly
                var fileExtension = Path.GetExtension(file.FileName).ToLowerInvariant();
                
                if (fileExtension == ".csv")
                {
                    await ProcessCsvFileAsync(tempFilePath, result);
                }
                else if (fileExtension == ".xlsx" || fileExtension == ".xls")
                {
                    await ProcessExcelFileAsync(tempFilePath, result);
                }
                else
                {
                    result.ValidationErrors.Add("Unsupported file format. Please upload .xlsx, .xls, or .csv files.");
                    result.Status = "error";
                }

                // Calculate quality score
                result.QualityScore = CalculateQualityScore(result);

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error analyzing file {FileName}", file.FileName);
                
                // Clean up temp file on error
                try
                {
                    if (File.Exists(tempFilePath))
                    {
                        File.Delete(tempFilePath);
                    }
                }
                catch (Exception deleteEx)
                {
                    _logger.LogWarning(deleteEx, "Failed to delete temp file {TempFilePath} after error", tempFilePath);
                }
                
                return new FileAnalysisResult
                {
                    FileId = fileId,
                    FileName = file.FileName,
                    FileSize = file.Length,
                    FileType = file.ContentType,
                    LastModified = DateTime.Now,
                    Status = "error",
                    ValidationErrors = new List<string> { $"Error processing file: {ex.Message}" }
                };
            }
        }

        private Task ProcessExcelFileAsync(string filePath, FileAnalysisResult result)
        {
            return Task.Run(() =>
            {
                try
                {
                    using var workbook = new XLWorkbook(filePath);
                    
                    _logger.LogInformation("Processing Excel file with {WorksheetCount} worksheets", workbook.Worksheets.Count);
                    
                                    var skippedWorksheets = new List<string>();
                
                foreach (var worksheet in workbook.Worksheets)
                {
                    var worksheetInfo = AnalyzeWorksheet(worksheet, result.FileId);
                    
                    if (worksheetInfo != null)
                    {
                        result.Worksheets.Add(worksheetInfo);
                        
                        _logger.LogInformation("Analyzed worksheet '{WorksheetName}' with {RowCount} rows and {ColumnCount} columns", 
                            worksheetInfo.Name, worksheetInfo.RowCount, worksheetInfo.ColumnCount);
                    }
                    else
                    {
                        skippedWorksheets.Add(worksheet.Name);
                        _logger.LogWarning("Skipped worksheet '{WorksheetName}' - no data found after headers", worksheet.Name);
                    }
                }
                
                // Add warnings for skipped worksheets
                if (skippedWorksheets.Any())
                {
                    result.ValidationWarnings.Add($"Skipped {skippedWorksheets.Count} worksheet(s) with no data: {string.Join(", ", skippedWorksheets)}");
                }

                    // If we have worksheets, select the first one by default and get its headers
                    if (result.Worksheets.Any())
                    {
                        result.Worksheets.First().Selected = true;
                        var selectedWorksheet = workbook.Worksheets.First();
                        result.Headers = GetWorksheetHeaders(selectedWorksheet);
                        
                        _logger.LogInformation("Selected first worksheet '{WorksheetName}' with {HeaderCount} headers", 
                            selectedWorksheet.Name, result.Headers.Count);
                    }

                    if (!result.Worksheets.Any())
                    {
                        result.ValidationErrors.Add("No valid worksheets found in the Excel file.");
                        _logger.LogWarning("No worksheets found in Excel file {FilePath}", filePath);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error processing Excel file {FilePath}", filePath);
                    result.ValidationErrors.Add($"Error reading Excel file: {ex.Message}");
                    result.Status = "error";
                }
            });
        }

        private async Task ProcessCsvFileAsync(string filePath, FileAnalysisResult result)
        {
            try
            {
                var lines = await File.ReadAllLinesAsync(filePath);
                if (lines.Length == 0)
                {
                    result.ValidationErrors.Add("CSV file is empty.");
                    return;
                }

                var worksheetInfo = new WorksheetInfo
                {
                    Id = "csv_main",
                    Name = "CSV Data",
                    RowCount = lines.Length, // Total rows including header
                    HasHeaders = true,
                    Selected = true
                };

                // Parse headers from first line
                var headerLine = lines[0];
                var headers = headerLine.Split(',').Select(h => h.Trim('"', ' ')).ToList();
                worksheetInfo.DetectedHeaders = headers;
                worksheetInfo.ColumnCount = headers.Count;

                // Get first data row preview
                if (lines.Length > 1)
                {
                    var previewValues = lines[1].Split(',').Select(cell => cell.Trim('"', ' ')).Take(5);
                    worksheetInfo.FirstDataRowPreview = string.Join(" | ", previewValues);
                }

                result.Worksheets.Add(worksheetInfo);
                result.Headers = MapCsvHeaders(headers);

                _logger.LogInformation("Successfully processed CSV file with {WorksheetCount} worksheet(s) and {HeaderCount} headers", 
                    result.Worksheets.Count, result.Headers.Count);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing CSV file {FilePath}", filePath);
                result.ValidationErrors.Add($"Error reading CSV file: {ex.Message}");
                result.Status = "error";
            }
        }

        private WorksheetInfo? AnalyzeWorksheet(IXLWorksheet worksheet, string fileId)
        {
            var worksheetInfo = new WorksheetInfo
            {
                Id = $"{fileId}_{worksheet.Name}",
                Name = worksheet.Name,
                HasHeaders = true // Assume first row is always headers as per user requirement
            };

            try
            {
                // Get used range
                var usedRange = worksheet.RangeUsed();
                if (usedRange == null)
                {
                    worksheetInfo.RowCount = 0;
                    worksheetInfo.ColumnCount = 0;
                    return worksheetInfo;
                }

                worksheetInfo.RowCount = usedRange.RowCount();
                
                // Check if there's actual data after headers (row 2 and beyond)
                bool hasDataAfterHeaders = false;
                if (worksheetInfo.RowCount > 1)
                {
                    // Check row 2 for any non-empty data
                    for (int col = 1; col <= usedRange.ColumnCount(); col++)
                    {
                        var cellValue = worksheet.Cell(2, col).GetString().Trim();
                        if (!string.IsNullOrEmpty(cellValue) && !IsLikelyFilterValue(cellValue))
                        {
                            hasDataAfterHeaders = true;
                            break;
                        }
                    }
                }

                // If no data found after headers, mark this worksheet as having issues
                if (!hasDataAfterHeaders && worksheetInfo.RowCount > 1)
                {
                    _logger.LogWarning("Worksheet '{WorksheetName}' has headers but no data in row 2", worksheet.Name);
                    // Don't return this worksheet - it will be filtered out
                    return null;
                }

                // Only count columns that have data (not just empty columns)
                var columnsWithData = GetColumnsWithData(worksheet, usedRange);
                worksheetInfo.ColumnCount = columnsWithData.Count;

                // Extract headers only from columns that have data
                if (worksheetInfo.RowCount > 0)
                {
                    foreach (var col in columnsWithData)
                    {
                        var cellValue = worksheet.Cell(1, col).GetString().Trim();
                        worksheetInfo.DetectedHeaders.Add(!string.IsNullOrEmpty(cellValue) ? cellValue : $"Column_{col}");
                    }
                }

                // Get first data row preview (row 2 since row 1 is headers)
                if (worksheetInfo.RowCount > 1 && hasDataAfterHeaders)
                {
                    var previewValues = new List<string>();
                    
                    foreach (var col in columnsWithData.Take(5))
                    {
                        var cellValue = worksheet.Cell(2, col).GetString().Trim();
                        previewValues.Add(string.IsNullOrEmpty(cellValue) ? "[empty]" : cellValue);
                    }
                    
                    worksheetInfo.FirstDataRowPreview = string.Join(" | ", previewValues);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error analyzing worksheet {WorksheetName}", worksheet.Name);
            }

            return worksheetInfo;
        }

        private List<int> GetColumnsWithData(IXLWorksheet worksheet, IXLRange usedRange)
        {
            var columnsWithData = new List<int>();
            
            try
            {
                for (int col = 1; col <= usedRange.ColumnCount(); col++)
                {
                    bool hasData = false;
                    
                    // Check if this column has any data from row 2 onwards (skip header row)
                    for (int row = 2; row <= usedRange.RowCount(); row++)
                    {
                        var cellValue = worksheet.Cell(row, col).GetString().Trim();
                        if (!string.IsNullOrEmpty(cellValue) && !IsLikelyFilterValue(cellValue))
                        {
                            hasData = true;
                            break;
                        }
                    }
                    
                    // Also check if the header itself has content (don't skip columns with headers but no data yet)
                    if (!hasData)
                    {
                        var headerValue = worksheet.Cell(1, col).GetString().Trim();
                        if (!string.IsNullOrEmpty(headerValue))
                        {
                            // Include columns with headers even if no data yet, but log it
                            hasData = true;
                            _logger.LogDebug("Column {Column} has header '{Header}' but no data", col, headerValue);
                        }
                    }
                    
                    if (hasData)
                    {
                        columnsWithData.Add(col);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting columns with data for worksheet {WorksheetName}", worksheet.Name);
                // Fallback: return all columns
                for (int col = 1; col <= usedRange.ColumnCount(); col++)
                {
                    columnsWithData.Add(col);
                }
            }
            
            return columnsWithData;
        }

        private List<HeaderMapping> GetWorksheetHeaders(IXLWorksheet worksheet)
        {
            var headers = new List<HeaderMapping>();
            
            try
            {
                var usedRange = worksheet.RangeUsed();
                if (usedRange == null || usedRange.RowCount() == 0)
                {
                    return headers;
                }

                // Get only columns that have data
                var columnsWithData = GetColumnsWithData(worksheet, usedRange);
                
                // Read headers only from columns that have data
                foreach (var col in columnsWithData)
                {
                    var headerValue = worksheet.Cell(1, col).GetString().Trim();
                    var displayName = !string.IsNullOrEmpty(headerValue) ? headerValue : $"Column_{col}";
                    
                    // Get sample data from the next few rows (skip potential filter rows)
                    var sampleValues = new List<string>();
                    int samplesFound = 0;
                    
                    // Start from row 2 and go deeper to find actual data (not filters)
                    for (int row = 2; row <= usedRange.RowCount() && samplesFound < 3; row++)
                    {
                        var cellValue = worksheet.Cell(row, col).GetString().Trim();
                        
                        // Skip cells that look like filter dropdowns or empty cells
                        if (!string.IsNullOrEmpty(cellValue) && !IsLikelyFilterValue(cellValue))
                        {
                            sampleValues.Add(cellValue);
                            samplesFound++;
                        }
                    }

                    var header = new HeaderMapping
                    {
                        Id = $"header_{col}",
                        DetectedName = displayName,
                        StandardName = displayName, // Keep the same name since we're not doing matching
                        DataType = DetermineDataType(sampleValues),
                        Selected = true, // Select all headers by default
                        IsRequired = false, // No predefined required fields
                        MatchConfidence = 1.0, // 100% since we're reading directly
                        Column = GetColumnLetter(col),
                        SampleData = string.Join(", ", sampleValues.Take(3))
                    };

                    headers.Add(header);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting worksheet headers for {WorksheetName}", worksheet.Name);
            }

            return headers;
        }

        private bool IsLikelyFilterValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return true;
                
            // Common filter dropdown indicators
            var filterIndicators = new[]
            {
                "(All)",
                "(Select All)",
                "(Multiple Items)",
                "Select...",
                "Choose...",
                "Filter...",
                "---",
                "...",
                "All"
            };
            
            // Check if the value matches common filter patterns
            foreach (var indicator in filterIndicators)
            {
                if (value.Equals(indicator, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            
            // Check if value is very generic (single letters, numbers, or short generic words)
            if (value.Length == 1 || (value.Length <= 3 && Regex.IsMatch(value, @"^[A-Za-z0-9]+$")))
                return true;
                
            return false;
        }

        private string DetermineDataType(List<string> sampleValues)
        {
            if (!sampleValues.Any()) return "Text";

            var dateCount = 0;
            var numberCount = 0;
            var boolCount = 0;

            foreach (var value in sampleValues)
            {
                if (DateTime.TryParse(value, out _))
                    dateCount++;
                else if (double.TryParse(value, out _))
                    numberCount++;
                else if (bool.TryParse(value, out _))
                    boolCount++;
            }

            var total = sampleValues.Count;
            if (dateCount > total * 0.6) return "Date";
            if (numberCount > total * 0.6) return "Numeric";
            if (boolCount > total * 0.6) return "Boolean";
            
            return "Text";
        }

        public async Task<DataPreviewResult> GetDataPreviewAsync(string filePath, string worksheetName, int maxRows = 10)
        {
            var result = new DataPreviewResult
            {
                WorksheetId = worksheetName
            };

            try
            {
                var fullPath = Path.Combine(_tempFilePath, filePath);
                
                if (Path.GetExtension(filePath).ToLowerInvariant() == ".csv")
                {
                    return await GetCsvPreviewAsync(fullPath, maxRows);
                }

                using var workbook = new XLWorkbook(fullPath);
                var worksheet = workbook.Worksheets.FirstOrDefault(w => w.Name == worksheetName);
                
                if (worksheet == null)
                {
                    return result;
                }

                var usedRange = worksheet.RangeUsed();
                if (usedRange == null)
                {
                    return result;
                }

                result.TotalRows = usedRange.RowCount();
                
                // Extract headers
                var headerRow = worksheet.Row(1);
                foreach (var cell in headerRow.CellsUsed())
                {
                    result.Headers.Add(cell.GetString().Trim());
                }

                // Extract data rows
                var startRow = 2; // Skip header
                var endRow = Math.Min(startRow + maxRows - 1, result.TotalRows);
                
                for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                {
                    var row = worksheet.Row(rowIndex);
                    var rowData = new List<string>();
                    
                    for (int colIndex = 1; colIndex <= result.Headers.Count; colIndex++)
                    {
                        var cellValue = worksheet.Cell(rowIndex, colIndex).GetString().Trim();
                        rowData.Add(cellValue);
                    }
                    
                    result.Rows.Add(rowData);
                }

                result.HasMoreData = result.TotalRows > maxRows + 1; // +1 for header
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting data preview for {FilePath}, worksheet {WorksheetName}", filePath, worksheetName);
            }

            return result;
        }

        private async Task<DataPreviewResult> GetCsvPreviewAsync(string filePath, int maxRows)
        {
            var result = new DataPreviewResult
            {
                WorksheetId = "csv_main"
            };

            try
            {
                var lines = await File.ReadAllLinesAsync(filePath);
                result.TotalRows = lines.Length - 1; // Excluding header

                if (lines.Length > 0)
                {
                    // Parse headers
                    var headerLine = lines[0];
                    result.Headers = headerLine.Split(',').Select(h => h.Trim('"', ' ')).ToList();

                    // Parse data rows
                    var dataLines = lines.Skip(1).Take(maxRows);
                    foreach (var line in dataLines)
                    {
                        var rowData = line.Split(',').Select(cell => cell.Trim('"', ' ')).ToList();
                        result.Rows.Add(rowData);
                    }

                    result.HasMoreData = lines.Length > maxRows + 1;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting CSV preview for {FilePath}", filePath);
            }

            return result;
        }

        public async Task<List<HeaderMapping>> DetectAndMapHeadersAsync(string filePath, string worksheetName)
        {
            try
            {
                var fullPath = Path.Combine(_tempFilePath, filePath);
                
                if (Path.GetExtension(filePath).ToLowerInvariant() == ".csv")
                {
                    var lines = await File.ReadAllLinesAsync(fullPath);
                    if (lines.Length > 0)
                    {
                        var headers = lines[0].Split(',').Select(h => h.Trim('"', ' ')).ToList();
                        return MapCsvHeaders(headers);
                    }
                    return new List<HeaderMapping>();
                }

                using var workbook = new XLWorkbook(fullPath);
                var worksheet = workbook.Worksheets.FirstOrDefault(w => w.Name == worksheetName);
                
                if (worksheet == null)
                {
                    return new List<HeaderMapping>();
                }

                return GetWorksheetHeaders(worksheet);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error detecting headers for {FilePath}, worksheet {WorksheetName}", filePath, worksheetName);
                return new List<HeaderMapping>();
            }
        }



        private List<HeaderMapping> MapCsvHeaders(List<string> headers)
        {
            var headerMappings = new List<HeaderMapping>();
            
            for (int i = 0; i < headers.Count; i++)
            {
                var detectedHeader = headers[i];
                var displayName = !string.IsNullOrEmpty(detectedHeader) ? detectedHeader : $"Column_{i + 1}";
                
                var mapping = new HeaderMapping
                {
                    Id = $"header_{i + 1}",
                    DetectedName = displayName,
                    StandardName = displayName,
                    DataType = "Text", // Default for CSV, could be enhanced with sample data analysis
                    Selected = true,
                    IsRequired = false,
                    MatchConfidence = 1.0,
                    Column = GetColumnLetter(i + 1),
                    SampleData = "" // Could be enhanced to read sample data
                };
                
                headerMappings.Add(mapping);
            }

            return headerMappings;
        }



        private string GetColumnLetter(int columnIndex)
        {
            string columnName = "";
            while (columnIndex > 0)
            {
                columnIndex--;
                columnName = (char)('A' + (columnIndex % 26)) + columnName;
                columnIndex /= 26;
            }
            return columnName;
        }

        private int CalculateQualityScore(FileAnalysisResult result)
        {
            if (result.ValidationErrors.Any())
                return 0;

            var score = 100;

            // Deduct points for warnings
            score -= result.ValidationWarnings.Count * 5;

            // Deduct points if no worksheets found
            if (!result.Worksheets.Any())
            {
                score -= 30;
            }

            // Deduct points if no headers detected
            if (!result.Headers.Any())
            {
                score -= 20;
            }

            return Math.Max(0, Math.Min(100, score));
        }

        public void CleanupOldTempFiles(int maxAgeHours = 24)
        {
            try
            {
                if (!Directory.Exists(_tempFilePath))
                    return;

                var cutoffTime = DateTime.Now.AddHours(-maxAgeHours);
                var tempFiles = Directory.GetFiles(_tempFilePath);

                foreach (var file in tempFiles)
                {
                    try
                    {
                        var fileInfo = new FileInfo(file);
                        if (fileInfo.CreationTime < cutoffTime)
                        {
                            File.Delete(file);
                            _logger.LogInformation("Deleted old temp file: {FileName}", fileInfo.Name);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Failed to delete temp file: {FileName}", file);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error during temp file cleanup");
            }
        }
    }
} 