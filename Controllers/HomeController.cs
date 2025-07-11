using System.Diagnostics;
using ExcelRefinery.Models;
using ExcelRefinery.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Antiforgery;
using Microsoft.AspNetCore.Http;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System;
using System.Collections.Generic;

namespace ExcelRefinery.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IExcelProcessingService _excelProcessingService;
        private readonly IWebHostEnvironment _hostingEnvironment;


        public HomeController(ILogger<HomeController> logger, IExcelProcessingService excelProcessingService, IWebHostEnvironment hostingEnvironment)
        {
            _logger = logger;
            _excelProcessingService = excelProcessingService;
            _hostingEnvironment = hostingEnvironment;
        }

        public IActionResult Index()
        {
            // Check if user is authenticated
            if (User.Identity?.IsAuthenticated == true)
            {
                // Redirect authenticated users to dashboard
                return View("Dashboard");
            }
            
            // Show welcome/landing page for non-authenticated users
            return View("Welcome");
        }

        public IActionResult Privacy()
        {
            return View();
        }

        public IActionResult StylingTemplate()
        {
            return View();
        }

        [Authorize]
        public IActionResult Upload()
        {
            return View();
        }

        [Authorize]
        public IActionResult ViewReports()
        {
            return View();
        }

        [HttpPost]
        [Authorize]
        public async Task<IActionResult> UploadFiles(IFormFileCollection files)
        {
            try
            {
                var results = new List<FileAnalysisResult>();

                foreach (var file in files)
                {
                    if (file.Length > 0)
                    {
                        var result = await _excelProcessingService.AnalyzeFileAsync(file);
                        results.Add(result);
                    }
                }

                return Json(new { success = true, files = results });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error uploading files");
                return Json(new { success = false, message = "Error processing files. Please try again." });
            }
        }

        [HttpGet]
        [Authorize]
        public async Task<IActionResult> GetDataPreview(string fileId, string worksheetName, int maxRows = 10)
        {
            try
            {
                // Find the uploaded file by fileId
                var uploadedFiles = Directory.GetFiles(Path.Combine("wwwroot", "temp"), $"{fileId}_*");
                if (!uploadedFiles.Any())
                {
                    return Json(new { success = false, message = "File not found." });
                }

                var filePath = Path.GetFileName(uploadedFiles.First());
                var result = await _excelProcessingService.GetDataPreviewAsync(filePath, worksheetName, maxRows);
                return Json(new { success = true, data = result });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting data preview for file {FileId}, worksheet {WorksheetName}", fileId, worksheetName);
                return Json(new { success = false, message = "Error loading data preview." });
            }
        }

        [HttpGet]
        [Authorize]
        public async Task<IActionResult> GetWorksheetHeaders(string fileId, string worksheetName)
        {
            try
            {
                // Find the uploaded file by fileId
                var uploadedFiles = Directory.GetFiles(Path.Combine("wwwroot", "temp"), $"{fileId}_*");
                if (!uploadedFiles.Any())
                {
                    return Json(new { success = false, message = "File not found." });
                }

                var filePath = Path.GetFileName(uploadedFiles.First());
                var headers = await _excelProcessingService.DetectAndMapHeadersAsync(filePath, worksheetName);
                return Json(new { success = true, headers = headers });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting headers for file {FileId}, worksheet {WorksheetName}", fileId, worksheetName);
                return Json(new { success = false, message = "Error loading worksheet headers." });
            }
        }

        [HttpPost]
        [Authorize]
        public async Task<IActionResult> CheckWorksheetIntegrity([FromBody] List<WorksheetComparisonRequest> requests)
        {
            try
            {
                if (requests == null || requests.Count == 0)
                {
                    return Json(new { success = false, message = "Please select worksheets from at least 2 files to check data integrity." });
                }

                _logger.LogInformation("Starting worksheet integrity check for {RequestCount} comparison requests", requests.Count);

                var integrityResults = await _excelProcessingService.CheckWorksheetIntegrityAsync(requests);
                
                _logger.LogInformation("Worksheet integrity check completed for {ComparisonCount} worksheet comparisons by user request", integrityResults.Count);
                return Json(new { success = true, results = integrityResults });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error checking worksheet integrity");
                return Json(new { success = false, message = "Error checking worksheet integrity. Please try again." });
            }
        }

        [HttpPost]
        [Authorize]
        public IActionResult ClearProcessedFileCache()
        {
            try
            {
                _excelProcessingService.ClearProcessedFileCache();
                _logger.LogInformation("Processed file cache cleared by user request");
                return Json(new { success = true, message = "Processed file cache cleared successfully." });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error clearing processed file cache");
                return Json(new { success = false, message = "Error clearing processed file cache." });
            }
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        #region API Endpoints for ViewReports

        [HttpPost("api/reports/upload")]
        [RequestSizeLimit(52428800)] // 50 MB limit
        [IgnoreAntiforgeryToken]
        public async Task<IActionResult> UploadFile(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest(new { message = "No file was received for processing." });
            }

            try
            {
                var result = await _excelProcessingService.AnalyzeFileAsync(file);
                if (result.Status == "error")
                {
                    return BadRequest(new { message = "Failed to analyze file.", errors = result.ValidationErrors });
                }
                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "File upload failed for {FileName}", file.FileName);
                return StatusCode(500, new { message = "An unexpected error occurred during file upload." });
            }
        }

        [HttpGet("api/reports/files")]
        [IgnoreAntiforgeryToken]
        public async Task<IActionResult> GetFiles()
        {
            var files = await _excelProcessingService.GetCachedFilesAsync();
            var fileList = files.Select(f => new { f.FileId, f.FileName }).ToList();
            return Ok(fileList);
        }

        [HttpGet("api/reports/worksheets/{fileId}")]
        [IgnoreAntiforgeryToken]
        public async Task<IActionResult> GetWorksheets(string fileId)
        {
            var file = (await _excelProcessingService.GetCachedFilesAsync()).FirstOrDefault(f => f.FileId == fileId);
            if (file == null)
            {
                return NotFound();
            }
            var worksheetNames = file.Worksheets.Select(w => w.Name).ToList();
            return Ok(worksheetNames);
        }

        [HttpGet("api/reports/preview/{fileId}/{worksheetName}")]
        [IgnoreAntiforgeryToken]
        public async Task<IActionResult> GetDataPreview(string fileId, string worksheetName)
        {
            var fileCache = (await _excelProcessingService.GetCachedFilesAsync()).FirstOrDefault(f => f.FileId == fileId);
            if (fileCache == null)
            {
                return NotFound("File not found in cache.");
            }
            
            var tempFileName = $"{fileId}_{Path.GetFileName(fileCache.FileName)}";
            var result = await _excelProcessingService.GetDataPreviewAsync(tempFileName, worksheetName);
            return Ok(result);
        }

        #endregion
    }
}
