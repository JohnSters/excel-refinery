/**
 * WorksheetComparisonService.cs
 * Provides advanced worksheet comparison capabilities for Excel refinery data
 * Handles position-independent header and row matching with configurable thresholds
 * Author: ExcelRefinery System
 */

using ExcelRefinery.Models;
using System.Text;
using System.Security.Cryptography;
using System.Text.RegularExpressions;

namespace ExcelRefinery.Services
{
    public interface IWorksheetComparisonService
    {
        Task<ComparisonResult> CompareWorksheetsAsync(NormalizedWorksheetData worksheet1, NormalizedWorksheetData worksheet2, double matchThreshold = 0.90);
        Task<List<WorksheetIntegrityComparison>> CompareWorksheetsBetweenFilesAsync(List<ProcessedFileCache> files, List<WorksheetComparisonRequest> requests);
        double CalculateRowSimilarity(Dictionary<string, string> row1, Dictionary<string, string> row2, Dictionary<string, string> headerMapping);
        List<HeaderMatchInfo> MatchHeaders(List<string> headers1, List<string> headers2);
        ComparisonStatus DetermineComparisonStatus(double similarityScore);
        SimilarityLevel DetermineSimilarityLevel(double similarityScore);
    }

    public class WorksheetComparisonService : IWorksheetComparisonService
    {
        private readonly ILogger<WorksheetComparisonService> _logger;
        
        // Configurable similarity thresholds
        private const double EXACT_MATCH_THRESHOLD = 1.0;          // 100%
        private const double NEAR_IDENTICAL_THRESHOLD = 0.98;      // 98%
        private const double HIGH_SIMILARITY_THRESHOLD = 0.90;     // 90%
        private const double MODERATE_SIMILARITY_THRESHOLD = 0.80;  // 80%
        private const double LOW_SIMILARITY_THRESHOLD = 0.50;      // 50%
        
        // Header matching thresholds
        private const double HEADER_EXACT_MATCH = 1.0;
        private const double HEADER_NORMALIZED_MATCH = 0.95;
        private const double HEADER_FUZZY_MATCH = 0.85;

        public WorksheetComparisonService(ILogger<WorksheetComparisonService> logger)
        {
            _logger = logger;
        }

        public async Task<ComparisonResult> CompareWorksheetsAsync(NormalizedWorksheetData worksheet1, NormalizedWorksheetData worksheet2, double matchThreshold = 0.90)
        {
            _logger.LogInformation("=== Starting Enhanced Worksheet Comparison ===");
            _logger.LogInformation("Worksheet 1: '{Name}' - {HeaderCount} headers, {RowCount} rows", 
                worksheet1.Name, worksheet1.Headers.Count, worksheet1.Rows.Count);
            _logger.LogInformation("Worksheet 2: '{Name}' - {HeaderCount} headers, {RowCount} rows", 
                worksheet2.Name, worksheet2.Headers.Count, worksheet2.Rows.Count);

            var comparison = new ComparisonResult
            {
                TotalRows = Math.Max(worksheet1.Rows.Count, worksheet2.Rows.Count)
            };

            try
            {
                // Step 1: Advanced header matching
                var headerMatches = MatchHeaders(worksheet1.Headers, worksheet2.Headers);
                comparison.HeaderMatches = headerMatches;
                
                var headerMapping = CreateOptimizedHeaderMapping(headerMatches);
                comparison.HeadersMatch = headerMatches.Count >= Math.Min(worksheet1.Headers.Count, worksheet2.Headers.Count) * 0.8; // 80% of headers must match
                
                _logger.LogInformation("Header matching complete: {MatchCount} matches out of {TotalHeaders} headers", 
                    headerMatches.Count, Math.Max(worksheet1.Headers.Count, worksheet2.Headers.Count));

                // Step 2: Identify missing and extra headers
                var mappedSourceHeaders = headerMatches.Select(h => h.SourceHeader).ToHashSet();
                var mappedTargetHeaders = headerMatches.Select(h => h.TargetHeader).ToHashSet();
                
                comparison.MissingHeaders = worksheet1.Headers.Where(h => !mappedSourceHeaders.Contains(h)).ToList();
                comparison.ExtraHeaders = worksheet2.Headers.Where(h => !mappedTargetHeaders.Contains(h)).ToList();
                comparison.MatchedHeaders = mappedSourceHeaders.ToList();

                if (!comparison.HeadersMatch)
                {
                    _logger.LogWarning("Header structure mismatch: {MissingCount} missing, {ExtraCount} extra headers", 
                        comparison.MissingHeaders.Count, comparison.ExtraHeaders.Count);
                    comparison.SummaryMessage = $"Header mismatch: {comparison.MissingHeaders.Count} missing, {comparison.ExtraHeaders.Count} extra headers";
                    comparison.SimilarityPercentage = 0.0;
                    comparison.Status = ComparisonStatus.Error;
                    comparison.SimilarityLevel = SimilarityLevel.Different;
                    return comparison;
                }

                // Step 3: Position-independent row comparison
                var (matchingRows, rowComparisons) = await CompareRowsAdvancedAsync(worksheet1.Rows, worksheet2.Rows, headerMapping, matchThreshold);
                
                comparison.MatchingRows = matchingRows;
                comparison.DifferentRows = comparison.TotalRows - matchingRows;
                
                // Step 4: Calculate similarity score
                double similarityScore = worksheet1.Rows.Count > 0 ? (double)matchingRows / worksheet1.Rows.Count : 1.0;
                comparison.SimilarityPercentage = similarityScore;
                comparison.SimilarityLevel = DetermineSimilarityLevel(similarityScore);
                comparison.Status = DetermineComparisonStatus(similarityScore);

                // Step 5: Generate detailed feedback
                comparison.DifferentRowSamples = GenerateRowDifferenceSamples(rowComparisons.Take(3).ToList());
                comparison.SummaryMessage = GenerateComparisonSummary(comparison);

                _logger.LogInformation("Worksheet comparison complete: {Similarity:P2} similarity, Status: {Status}, Level: {Level}", 
                    similarityScore, comparison.Status, comparison.SimilarityLevel);

                return comparison;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error during worksheet comparison");
                comparison.SummaryMessage = $"Comparison failed: {ex.Message}";
                comparison.Status = ComparisonStatus.Error;
                return comparison;
            }
        }

        public List<HeaderMatchInfo> MatchHeaders(List<string> headers1, List<string> headers2)
        {
            var matches = new List<HeaderMatchInfo>();
            var usedHeaders2 = new HashSet<string>();

            _logger.LogInformation("Matching headers: Set1=[{Headers1}], Set2=[{Headers2}]", 
                string.Join(", ", headers1), string.Join(", ", headers2));

            foreach (var header1 in headers1)
            {
                HeaderMatchInfo? bestMatch = null;
                double bestScore = 0.0;

                foreach (var header2 in headers2.Where(h => !usedHeaders2.Contains(h)))
                {
                    var matchInfo = CalculateHeaderMatch(header1, header2);
                    if (matchInfo.MatchConfidence > bestScore && matchInfo.MatchConfidence >= HEADER_FUZZY_MATCH)
                    {
                        bestScore = matchInfo.MatchConfidence;
                        bestMatch = matchInfo;
                    }
                }

                if (bestMatch != null)
                {
                    matches.Add(bestMatch);
                    usedHeaders2.Add(bestMatch.TargetHeader);
                    _logger.LogDebug("Header match: '{Source}' -> '{Target}' ({Confidence:P2}, {Reason})", 
                        bestMatch.SourceHeader, bestMatch.TargetHeader, bestMatch.MatchConfidence, bestMatch.MatchReason);
                }
                else
                {
                    _logger.LogDebug("No match found for header: '{Header}'", header1);
                }
            }

            _logger.LogInformation("Header matching complete: {MatchCount}/{TotalCount} headers matched", 
                matches.Count, headers1.Count);

            return matches;
        }

        private HeaderMatchInfo CalculateHeaderMatch(string header1, string header2)
        {
            // Exact match (highest priority)
            if (string.Equals(header1, header2, StringComparison.OrdinalIgnoreCase))
            {
                return new HeaderMatchInfo
                {
                    SourceHeader = header1,
                    TargetHeader = header2,
                    MatchConfidence = HEADER_EXACT_MATCH,
                    MatchReason = "exact"
                };
            }

            // Normalized match (handle spaces, underscores, case)
            var normalized1 = NormalizeHeaderName(header1);
            var normalized2 = NormalizeHeaderName(header2);
            
            if (normalized1 == normalized2)
            {
                return new HeaderMatchInfo
                {
                    SourceHeader = header1,
                    TargetHeader = header2,
                    MatchConfidence = HEADER_NORMALIZED_MATCH,
                    MatchReason = "normalized"
                };
            }

            // Fuzzy matching using Levenshtein distance
            var fuzzyScore = CalculateLevenshteinSimilarity(normalized1, normalized2);
            if (fuzzyScore >= HEADER_FUZZY_MATCH)
            {
                return new HeaderMatchInfo
                {
                    SourceHeader = header1,
                    TargetHeader = header2,
                    MatchConfidence = fuzzyScore,
                    MatchReason = "fuzzy"
                };
            }

            // No match
            return new HeaderMatchInfo
            {
                SourceHeader = header1,
                TargetHeader = header2,
                MatchConfidence = fuzzyScore,
                MatchReason = "no_match"
            };
        }

        private async Task<(int matchingRows, List<RowComparisonResult>)> CompareRowsAdvancedAsync(
            List<Dictionary<string, string>> rows1, 
            List<Dictionary<string, string>> rows2, 
            Dictionary<string, string> headerMapping, 
            double matchThreshold)
        {
            _logger.LogInformation("Starting advanced row comparison with {Rows1Count} vs {Rows2Count} rows, threshold: {Threshold:P2}", 
                rows1.Count, rows2.Count, matchThreshold);

            var rowComparisons = new List<RowComparisonResult>();
            var matchedRows2 = new HashSet<int>();
            int matchingRowCount = 0;

            // Create searchable index for rows2 for better performance
            var rows2Index = CreateRowIndex(rows2, headerMapping.Values.ToList());

            for (int i = 0; i < rows1.Count; i++)
            {
                var row1 = rows1[i];
                var bestMatch = FindBestRowMatch(row1, rows2, headerMapping, matchedRows2, matchThreshold, rows2Index);
                
                if (bestMatch != null)
                {
                    matchedRows2.Add(bestMatch.Row2Index);
                    matchingRowCount++;
                    
                    rowComparisons.Add(new RowComparisonResult
                    {
                        Row1Index = i,
                        Row2Index = bestMatch.Row2Index,
                        SimilarityScore = bestMatch.SimilarityScore,
                        IsMatch = true,
                        Differences = bestMatch.Differences
                    });
                }
                else
                {
                    rowComparisons.Add(new RowComparisonResult
                    {
                        Row1Index = i,
                        Row2Index = -1,
                        SimilarityScore = 0.0,
                        IsMatch = false,
                        Differences = new List<string> { "No matching row found" }
                    });
                }
            }

            _logger.LogInformation("Row comparison complete: {MatchingRows}/{TotalRows1} rows matched ({Percentage:P1})", 
                matchingRowCount, rows1.Count, rows1.Count > 0 ? (double)matchingRowCount / rows1.Count : 0);

            return (matchingRowCount, rowComparisons);
        }

        private Dictionary<string, List<int>> CreateRowIndex(List<Dictionary<string, string>> rows, List<string> keyHeaders)
        {
            var index = new Dictionary<string, List<int>>();
            
            for (int i = 0; i < rows.Count; i++)
            {
                var row = rows[i];
                foreach (var header in keyHeaders.Take(3)) // Index by first 3 key headers for performance
                {
                    var value = row.GetValueOrDefault(header, "").Trim().ToLowerInvariant();
                    if (!string.IsNullOrEmpty(value))
                    {
                        var key = $"{header}:{value}";
                        if (!index.ContainsKey(key))
                            index[key] = new List<int>();
                        index[key].Add(i);
                    }
                }
            }
            
            return index;
        }

        private RowMatchResult? FindBestRowMatch(
            Dictionary<string, string> row1, 
            List<Dictionary<string, string>> rows2, 
            Dictionary<string, string> headerMapping, 
            HashSet<int> usedRows2, 
            double threshold,
            Dictionary<string, List<int>> rows2Index)
        {
            var candidates = new List<int>();
            
            // First, try to find candidates using the index
            foreach (var mapping in headerMapping.Take(3)) // Use first 3 headers for candidate selection
            {
                var value1 = row1.GetValueOrDefault(mapping.Key, "").Trim().ToLowerInvariant();
                if (!string.IsNullOrEmpty(value1))
                {
                    var key = $"{mapping.Value}:{value1}";
                    if (rows2Index.ContainsKey(key))
                    {
                        candidates.AddRange(rows2Index[key].Where(idx => !usedRows2.Contains(idx)));
                    }
                }
            }

            // If no candidates found via index, check all unused rows (fallback)
            if (!candidates.Any())
            {
                candidates = Enumerable.Range(0, rows2.Count).Where(idx => !usedRows2.Contains(idx)).ToList();
            }

            // Remove duplicates and limit candidates for performance
            candidates = candidates.Distinct().Take(50).ToList();

            RowMatchResult? bestMatch = null;
            double bestScore = 0.0;

            foreach (var candidateIndex in candidates)
            {
                var row2 = rows2[candidateIndex];
                var similarity = CalculateRowSimilarity(row1, row2, headerMapping);
                
                if (similarity >= threshold && similarity > bestScore)
                {
                    bestScore = similarity;
                    bestMatch = new RowMatchResult
                    {
                        Row2Index = candidateIndex,
                        SimilarityScore = similarity,
                        Differences = GenerateRowDifferences(row1, row2, headerMapping)
                    };
                }
            }

            return bestMatch;
        }

        public double CalculateRowSimilarity(Dictionary<string, string> row1, Dictionary<string, string> row2, Dictionary<string, string> headerMapping)
        {
            if (!headerMapping.Any()) return 0.0;

            int totalFields = headerMapping.Count;
            int matchingFields = 0;
            double totalSimilarity = 0.0;

            foreach (var mapping in headerMapping)
            {
                var value1 = row1.GetValueOrDefault(mapping.Key, "").Trim();
                var value2 = row2.GetValueOrDefault(mapping.Value, "").Trim();

                var fieldSimilarity = CalculateFieldSimilarity(value1, value2);
                totalSimilarity += fieldSimilarity;

                if (fieldSimilarity >= 0.9) // 90% field-level threshold
                {
                    matchingFields++;
                }
            }

            // Average similarity across all fields
            return totalSimilarity / totalFields;
        }

        private double CalculateFieldSimilarity(string value1, string value2)
        {
            // Handle empty values
            if (string.IsNullOrEmpty(value1) && string.IsNullOrEmpty(value2)) return 1.0;
            if (string.IsNullOrEmpty(value1) || string.IsNullOrEmpty(value2)) return 0.0;

            // Exact match
            if (string.Equals(value1, value2, StringComparison.OrdinalIgnoreCase)) return 1.0;

            // Normalize and compare
            var normalized1 = NormalizeFieldValue(value1);
            var normalized2 = NormalizeFieldValue(value2);
            
            if (normalized1 == normalized2) return 0.95;

            // Use Levenshtein distance for fuzzy matching
            return CalculateLevenshteinSimilarity(normalized1, normalized2);
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

        private string NormalizeFieldValue(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";

            return value
                .ToLowerInvariant()
                .Trim()
                .Replace("  ", " ")  // Remove double spaces
                .Replace("\t", " ")  // Replace tabs with spaces
                .Replace("\n", " ")  // Replace newlines with spaces
                .Replace("\r", "");  // Remove carriage returns
        }

        private double CalculateLevenshteinSimilarity(string str1, string str2)
        {
            if (str1 == str2) return 1.0;
            if (string.IsNullOrEmpty(str1) || string.IsNullOrEmpty(str2)) return 0.0;

            int maxLength = Math.Max(str1.Length, str2.Length);
            if (maxLength == 0) return 1.0;

            int distance = CalculateLevenshteinDistance(str1, str2);
            return 1.0 - (double)distance / maxLength;
        }

        private int CalculateLevenshteinDistance(string str1, string str2)
        {
            if (str1 == str2) return 0;
            if (str1.Length == 0) return str2.Length;
            if (str2.Length == 0) return str1.Length;

            int[,] matrix = new int[str1.Length + 1, str2.Length + 1];

            for (int i = 0; i <= str1.Length; i++) matrix[i, 0] = i;
            for (int j = 0; j <= str2.Length; j++) matrix[0, j] = j;

            for (int i = 1; i <= str1.Length; i++)
            {
                for (int j = 1; j <= str2.Length; j++)
                {
                    int cost = str1[i - 1] == str2[j - 1] ? 0 : 1;
                    matrix[i, j] = Math.Min(
                        Math.Min(matrix[i - 1, j] + 1, matrix[i, j - 1] + 1),
                        matrix[i - 1, j - 1] + cost);
                }
            }

            return matrix[str1.Length, str2.Length];
        }

        public ComparisonStatus DetermineComparisonStatus(double similarityScore)
        {
            return similarityScore switch
            {
                >= HIGH_SIMILARITY_THRESHOLD => ComparisonStatus.Success,  // 90%+ = Success (Green)
                >= LOW_SIMILARITY_THRESHOLD => ComparisonStatus.Warning,   // 50-89% = Warning (Yellow)
                _ => ComparisonStatus.Error                                 // <50% = Error (Red)
            };
        }

        public SimilarityLevel DetermineSimilarityLevel(double similarityScore)
        {
            return similarityScore switch
            {
                >= EXACT_MATCH_THRESHOLD => SimilarityLevel.ExactMatch,
                >= NEAR_IDENTICAL_THRESHOLD => SimilarityLevel.NearIdentical,
                >= HIGH_SIMILARITY_THRESHOLD => SimilarityLevel.HighSimilarity,
                >= MODERATE_SIMILARITY_THRESHOLD => SimilarityLevel.ModerateSimilarity,
                >= LOW_SIMILARITY_THRESHOLD => SimilarityLevel.LowSimilarity,
                _ => SimilarityLevel.Different
            };
        }

        public async Task<List<WorksheetIntegrityComparison>> CompareWorksheetsBetweenFilesAsync(
            List<ProcessedFileCache> files, 
            List<WorksheetComparisonRequest> requests)
        {
            var comparisons = new List<WorksheetIntegrityComparison>();

            _logger.LogInformation("=== Starting Multi-File Worksheet Comparison ===");
            _logger.LogInformation("Processing {RequestCount} comparison requests across {FileCount} files", 
                requests.Count, files.Count);

            foreach (var request in requests)
            {
                try
                {
                    var file1 = files.FirstOrDefault(f => f.FileId == request.File1Id);
                    var file2 = files.FirstOrDefault(f => f.FileId == request.File2Id);

                    if (file1 == null || file2 == null)
                    {
                        _logger.LogWarning("File not found for comparison request: File1={File1Id}, File2={File2Id}", 
                            request.File1Id, request.File2Id);
                        continue;
                    }

                    var worksheet1 = file1.Worksheets.FirstOrDefault(w => w.Name == request.File1WorksheetName);
                    var worksheet2 = file2.Worksheets.FirstOrDefault(w => w.Name == request.File2WorksheetName);

                    if (worksheet1 == null || worksheet2 == null)
                    {
                        _logger.LogWarning("Worksheet not found: WS1={WS1} in {File1}, WS2={WS2} in {File2}", 
                            request.File1WorksheetName, file1.FileName, request.File2WorksheetName, file2.FileName);
                        continue;
                    }

                    _logger.LogInformation("Comparing: {File1}[{WS1}] vs {File2}[{WS2}]", 
                        file1.FileName, worksheet1.Name, file2.FileName, worksheet2.Name);

                    var comparisonResult = await CompareWorksheetsAsync(worksheet1, worksheet2, request.MatchThreshold);

                    var comparison = new WorksheetIntegrityComparison
                    {
                        SourceWorksheetName = worksheet1.Name,
                        ComparedWithFileId = file2.FileId,
                        ComparedWithFileName = file2.FileName,
                        ComparedWithWorksheetName = worksheet2.Name,
                        SimilarityScore = comparisonResult.SimilarityPercentage,
                        SimilarityLevel = comparisonResult.SimilarityLevel,
                        Status = comparisonResult.Status,
                        DetailedComparison = comparisonResult,
                        SpecificDifferences = CreateSpecificDifferencesList(comparisonResult)
                    };

                    // Set UI properties based on status
                    SetComparisonUIProperties(comparison);

                    comparisons.Add(comparison);

                    _logger.LogInformation("Comparison complete: {Similarity:P2} similarity, Status: {Status}", 
                        comparison.SimilarityScore, comparison.Status);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error comparing worksheets for request: {Request}", request);
                }
            }

            _logger.LogInformation("Multi-file worksheet comparison complete: {ComparisonCount} comparisons processed", 
                comparisons.Count);

            return comparisons;
        }

        private Dictionary<string, string> CreateOptimizedHeaderMapping(List<HeaderMatchInfo> headerMatches)
        {
            return headerMatches
                .Where(h => h.MatchConfidence >= HEADER_FUZZY_MATCH)
                .ToDictionary(h => h.SourceHeader, h => h.TargetHeader);
        }

        private List<string> GenerateRowDifferenceSamples(List<RowComparisonResult> rowComparisons)
        {
            var samples = new List<string>();
            
            foreach (var comparison in rowComparisons.Where(c => !c.IsMatch).Take(3))
            {
                if (comparison.Row2Index == -1)
                {
                    samples.Add($"Row {comparison.Row1Index + 1}: No matching row found");
                }
                else
                {
                    var differences = string.Join(", ", comparison.Differences.Take(2));
                    samples.Add($"Row {comparison.Row1Index + 1} vs Row {comparison.Row2Index + 1}: {differences}");
                }
            }
            
            return samples;
        }

        private string GenerateComparisonSummary(ComparisonResult comparison)
        {
            var similarity = comparison.SimilarityPercentage;
            
            return comparison.Status switch
            {
                ComparisonStatus.Success when similarity >= EXACT_MATCH_THRESHOLD => 
                    "✅ Perfect match - All data is identical between worksheets",
                ComparisonStatus.Success when similarity >= NEAR_IDENTICAL_THRESHOLD => 
                    $"✅ Near perfect match - {comparison.SimilarityPercentage:P1} similarity, excellent data consistency",
                ComparisonStatus.Success => 
                    $"✅ Good data consistency - {comparison.SimilarityPercentage:P1} similarity, minor differences detected",
                ComparisonStatus.Warning => 
                    $"⚠️ Some differences found - {comparison.SimilarityPercentage:P1} similarity, review recommended",
                ComparisonStatus.Error => 
                    $"❌ Significant differences - {comparison.SimilarityPercentage:P1} similarity, data may not be consistent",
                _ => $"❓ Unknown comparison result - {comparison.SimilarityPercentage:P1} similarity"
            };
        }

        private List<string> CreateSpecificDifferencesList(ComparisonResult comparison)
        {
            var differences = new List<string>();

            if (!comparison.HeadersMatch)
            {
                if (comparison.MissingHeaders?.Any() == true)
                    differences.Add($"Missing columns: {string.Join(", ", comparison.MissingHeaders)}");
                if (comparison.ExtraHeaders?.Any() == true)
                    differences.Add($"Extra columns: {string.Join(", ", comparison.ExtraHeaders)}");
            }

            if (comparison.DifferentRows > 0)
            {
                differences.Add($"{comparison.DifferentRows} of {comparison.TotalRows} rows have differences");
                if (comparison.DifferentRowSamples?.Any() == true)
                    differences.AddRange(comparison.DifferentRowSamples.Take(3));
            }

            if (!differences.Any())
                differences.Add("All data matches exactly");

            return differences;
        }

        private void SetComparisonUIProperties(WorksheetIntegrityComparison comparison)
        {
            switch (comparison.Status)
            {
                case ComparisonStatus.Success:
                    comparison.StatusIcon = "check_circle";
                    comparison.StatusColor = "success";
                    if (comparison.SimilarityScore >= EXACT_MATCH_THRESHOLD)
                        comparison.StatusMessage = $"✅ Perfect Match ({comparison.SimilarityScore:P1})";
                    else if (comparison.SimilarityScore >= NEAR_IDENTICAL_THRESHOLD)
                        comparison.StatusMessage = $"✅ Near Perfect Match ({comparison.SimilarityScore:P1})";
                    else
                        comparison.StatusMessage = $"✅ Good Consistency ({comparison.SimilarityScore:P1})";
                    break;
                
                case ComparisonStatus.Warning:
                    comparison.StatusIcon = "warning";
                    comparison.StatusColor = "warning";
                    comparison.StatusMessage = $"⚠️ Some Differences ({comparison.SimilarityScore:P1})";
                    break;
                
                case ComparisonStatus.Error:
                    comparison.StatusIcon = "error";
                    comparison.StatusColor = "danger";
                    comparison.StatusMessage = $"❌ Significant Differences ({comparison.SimilarityScore:P1})";
                    break;
                
                default:
                    comparison.StatusIcon = "help";
                    comparison.StatusColor = "secondary";
                    comparison.StatusMessage = $"❓ Unknown Status ({comparison.SimilarityScore:P1})";
                    break;
            }
        }

        private List<string> GenerateRowDifferences(Dictionary<string, string> row1, Dictionary<string, string> row2, Dictionary<string, string> headerMapping)
        {
            var differences = new List<string>();
            
            foreach (var mapping in headerMapping.Take(3)) // Limit to first 3 differences for readability
            {
                var value1 = row1.GetValueOrDefault(mapping.Key, "").Trim();
                var value2 = row2.GetValueOrDefault(mapping.Value, "").Trim();
                
                if (!string.Equals(value1, value2, StringComparison.OrdinalIgnoreCase))
                {
                    differences.Add($"{mapping.Key}: '{value1}' vs '{value2}'");
                }
            }
            
            return differences;
        }

        // Helper classes for internal processing
        private class RowMatchResult
        {
            public int Row2Index { get; set; }
            public double SimilarityScore { get; set; }
            public List<string> Differences { get; set; } = new();
        }

        private class RowComparisonResult
        {
            public int Row1Index { get; set; }
            public int Row2Index { get; set; }
            public double SimilarityScore { get; set; }
            public bool IsMatch { get; set; }
            public List<string> Differences { get; set; } = new();
        }
    }
} 