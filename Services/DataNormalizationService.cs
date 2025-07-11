/**
 * DataNormalizationService.cs
 * Handles normalization of Excel and CSV data for comparison purposes
 * Converts raw file data into standardized format for analysis
 * Author: ExcelRefinery System
 */

using ClosedXML.Excel;
using ExcelRefinery.Models;
using System.Security.Cryptography;
using System.Text;

namespace ExcelRefinery.Services
{
    public interface IDataNormalizationService
    {
        Task<List<NormalizedWorksheetData>> NormalizeFileDataAsync(string filePath, FileAnalysisResult fileAnalysis);
        Task<NormalizedWorksheetData?> NormalizeWorksheetAsync(string filePath, string worksheetName);
        Task<string> CalculateFileHashAsync(string filePath);
        string CalculateDataHash(List<string> headers, List<Dictionary<string, string>> rows);
    }

    public class DataNormalizationService : IDataNormalizationService
    {
        private readonly ILogger<DataNormalizationService> _logger;
        private readonly string _tempFilePath;

        public DataNormalizationService(ILogger<DataNormalizationService> logger, IWebHostEnvironment environment)
        {
            _logger = logger;
            _tempFilePath = Path.Combine(environment.WebRootPath, "temp");
        }

        public async Task<List<NormalizedWorksheetData>> NormalizeFileDataAsync(string filePath, FileAnalysisResult fileAnalysis)
        {
            var normalizedWorksheets = new List<NormalizedWorksheetData>();
            
            try
            {
                var fileExtension = Path.GetExtension(filePath).ToLowerInvariant();
                var fullPath = Path.Combine(_tempFilePath, filePath);

                _logger.LogInformation("Starting data normalization for {FileName} ({Extension})", 
                    Path.GetFileName(filePath), fileExtension);

                if (fileExtension == ".csv")
                {
                    var csvData = await NormalizeCsvDataAsync(fullPath);
                    if (csvData != null)
                    {
                        normalizedWorksheets.Add(csvData);
                    }
                }
                else if (fileExtension == ".xlsx" || fileExtension == ".xls")
                {
                    var excelData = NormalizeExcelData(fullPath, fileAnalysis.Worksheets);
                    normalizedWorksheets.AddRange(excelData);
                }

                _logger.LogInformation("Data normalization complete: {WorksheetCount} worksheets normalized", 
                    normalizedWorksheets.Count);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error normalizing file data for {FilePath}", filePath);
            }
            
            return normalizedWorksheets;
        }

        public async Task<NormalizedWorksheetData?> NormalizeWorksheetAsync(string filePath, string worksheetName)
        {
            try
            {
                var fileExtension = Path.GetExtension(filePath).ToLowerInvariant();
                var fullPath = Path.Combine(_tempFilePath, filePath);

                if (fileExtension == ".csv")
                {
                    return await NormalizeCsvDataAsync(fullPath);
                }
                else if (fileExtension == ".xlsx" || fileExtension == ".xls")
                {
                    using var workbook = new XLWorkbook(fullPath);
                    var worksheet = workbook.Worksheets.FirstOrDefault(w => w.Name == worksheetName);
                    
                    if (worksheet != null)
                    {
                        return NormalizeExcelWorksheet(worksheet);
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error normalizing worksheet {WorksheetName} from {FilePath}", worksheetName, filePath);
                return null;
            }
        }

        private async Task<NormalizedWorksheetData?> NormalizeCsvDataAsync(string filePath)
        {
            try
            {
                var lines = await File.ReadAllLinesAsync(filePath);
                if (lines.Length < 2) // Need at least header and one data row
                {
                    _logger.LogWarning("CSV file has insufficient data rows: {LineCount}", lines.Length);
                    return null;
                }

                var headers = lines[0].Split(',').Select(h => h.Trim('"', ' ')).ToList();
                var rows = new List<Dictionary<string, string>>();

                _logger.LogInformation("Normalizing CSV with {HeaderCount} headers and {DataRowCount} data rows", 
                    headers.Count, lines.Length - 1);

                for (int i = 1; i < lines.Length; i++)
                {
                    var values = lines[i].Split(',').Select(v => v.Trim('"', ' ')).ToList();
                    var row = new Dictionary<string, string>();
                    
                    for (int j = 0; j < Math.Min(headers.Count, values.Count); j++)
                    {
                        row[headers[j]] = values[j];
                    }
                    
                    // Only add rows with actual data
                    if (row.Values.Any(v => !string.IsNullOrWhiteSpace(v)))
                    {
                        rows.Add(row);
                    }
                }

                return new NormalizedWorksheetData
                {
                    Name = "csv_main",
                    Headers = headers,
                    Rows = rows,
                    OriginalRowCount = lines.Length,
                    DataRowCount = rows.Count,
                    DataHash = CalculateDataHash(headers, rows)
                };
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error normalizing CSV data from {FilePath}", filePath);
                return null;
            }
        }

        private List<NormalizedWorksheetData> NormalizeExcelData(string filePath, List<WorksheetInfo> worksheetInfos)
        {
            var normalizedWorksheets = new List<NormalizedWorksheetData>();
            
            try
            {
                using var workbook = new XLWorkbook(filePath);
                
                _logger.LogInformation("Normalizing Excel file: {FilePath} with {WorksheetCount} worksheets", 
                    Path.GetFileName(filePath), workbook.Worksheets.Count);
                
                foreach (var worksheetInfo in worksheetInfos)
                {
                    var worksheet = workbook.Worksheets.FirstOrDefault(w => w.Name == worksheetInfo.Name);
                    if (worksheet != null)
                    {
                        var normalizedData = NormalizeExcelWorksheet(worksheet);
                        if (normalizedData != null)
                        {
                            normalizedWorksheets.Add(normalizedData);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error normalizing Excel data from {FilePath}", filePath);
            }
            
            return normalizedWorksheets;
        }

        private NormalizedWorksheetData? NormalizeExcelWorksheet(IXLWorksheet worksheet)
        {
            try
            {
                _logger.LogInformation("Normalizing worksheet: '{WorksheetName}'", worksheet.Name);
                
                var usedRange = worksheet.RangeUsed();
                if (usedRange == null)
                {
                    _logger.LogWarning("Worksheet '{WorksheetName}' has no used range - skipping", worksheet.Name);
                    return null;
                }
                
                if (usedRange.RowCount() < 2) // Need at least header and one data row
                {
                    _logger.LogWarning("Worksheet '{WorksheetName}' has insufficient rows ({RowCount}) - skipping", 
                        worksheet.Name, usedRange.RowCount());
                    return null;
                }

                // Get columns with data
                var columnsWithData = GetColumnsWithData(worksheet, usedRange);
                if (!columnsWithData.Any())
                {
                    _logger.LogWarning("Worksheet '{WorksheetName}' has no columns with data", worksheet.Name);
                    return null;
                }

                // Extract headers (with improved normalization)
                var headers = new List<string>();
                foreach (var col in columnsWithData)
                {
                    var headerValue = worksheet.Cell(1, col).GetString().Trim();
                    headers.Add(!string.IsNullOrEmpty(headerValue) ? headerValue : $"column_{col}");
                }

                // Extract data rows (starting from row 3 to skip headers and filters)
                var rows = new List<Dictionary<string, string>>();
                for (int rowIndex = 3; rowIndex <= usedRange.RowCount(); rowIndex++)
                {
                    var row = new Dictionary<string, string>();
                    bool hasData = false;
                    
                    for (int i = 0; i < columnsWithData.Count; i++)
                    {
                        var col = columnsWithData[i];
                        var cellValue = worksheet.Cell(rowIndex, col).GetString().Trim();
                        
                        if (!string.IsNullOrEmpty(cellValue) && !IsLikelyFilterValue(cellValue))
                        {
                            hasData = true;
                        }
                        
                        row[headers[i]] = cellValue;
                    }
                    
                    if (hasData)
                    {
                        rows.Add(row);
                    }
                }

                if (rows.Any())
                {
                    _logger.LogInformation("Normalized worksheet '{WorksheetName}': {HeaderCount} headers, {RowCount} data rows", 
                        worksheet.Name, headers.Count, rows.Count);
                    
                    // Log sample data for verification
                    if (rows.Count > 0)
                    {
                        var sampleRow = string.Join(", ", headers.Take(3).Select(h => 
                            $"{h}='{rows[0].GetValueOrDefault(h, "")}'"));
                        _logger.LogDebug("Sample data from '{WorksheetName}': {SampleData}", worksheet.Name, sampleRow);
                    }

                    return new NormalizedWorksheetData
                    {
                        Name = worksheet.Name,
                        Headers = headers,
                        Rows = rows,
                        OriginalRowCount = usedRange.RowCount(),
                        DataRowCount = rows.Count,
                        DataHash = CalculateDataHash(headers, rows)
                    };
                }
                else
                {
                    _logger.LogWarning("Skipping worksheet '{WorksheetName}' - no data rows found", worksheet.Name);
                    return null;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error normalizing worksheet {WorksheetName}", worksheet.Name);
                return null;
            }
        }

        private List<int> GetColumnsWithData(IXLWorksheet worksheet, IXLRange usedRange)
        {
            var columnsWithData = new List<int>();
            
            try
            {
                for (int col = 1; col <= usedRange.ColumnCount(); col++)
                {
                    bool hasData = false;
                    
                    // Check if this column has any data from row 3 onwards (skip header row 1 and filter row 2)
                    for (int row = 3; row <= usedRange.RowCount(); row++)
                    {
                        var cellValue = worksheet.Cell(row, col).GetString().Trim();
                        if (!string.IsNullOrEmpty(cellValue) && !IsLikelyFilterValue(cellValue))
                        {
                            hasData = true;
                            break;
                        }
                    }
                    
                    // Also check if the header itself has content
                    if (!hasData)
                    {
                        var headerValue = worksheet.Cell(1, col).GetString().Trim();
                        if (!string.IsNullOrEmpty(headerValue))
                        {
                            hasData = true;
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
            if (value.Length == 1 || (value.Length <= 3 && System.Text.RegularExpressions.Regex.IsMatch(value, @"^[A-Za-z0-9]+$")))
                return true;
                
            return false;
        }

        public string CalculateDataHash(List<string> headers, List<Dictionary<string, string>> rows)
        {
            try
            {
                var sb = new StringBuilder();
                
                // Add headers to hash (preserve order)
                foreach (var header in headers)
                {
                    sb.Append(NormalizeHeaderName(header));
                    sb.Append("|");
                }
                
                // Add rows to hash in original order (position-dependent)
                foreach (var row in rows)
                {
                    foreach (var header in headers)
                    {
                        sb.Append(row.GetValueOrDefault(header, "").Trim());
                        sb.Append("|");
                    }
                    sb.Append("\n");
                }
                
                using var md5 = MD5.Create();
                var hashBytes = md5.ComputeHash(Encoding.UTF8.GetBytes(sb.ToString()));
                return BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error calculating data hash");
                return Guid.NewGuid().ToString();
            }
        }

        public async Task<string> CalculateFileHashAsync(string filePath)
        {
            try
            {
                var fullPath = Path.Combine(_tempFilePath, filePath);
                using var md5 = MD5.Create();
                using var stream = File.OpenRead(fullPath);
                var hashBytes = await md5.ComputeHashAsync(stream);
                return BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error calculating file hash for {FilePath}", filePath);
                return Guid.NewGuid().ToString();
            }
        }

        private string NormalizeHeaderName(string headerName)
        {
            if (string.IsNullOrEmpty(headerName)) return "";

            return headerName
                .ToLowerInvariant()
                .Replace(" ", "_")
                .Replace("-", "_")
                .Replace(".", "_")
                .Replace("(", "")
                .Replace(")", "")
                .Trim('_');
        }
    }
} 