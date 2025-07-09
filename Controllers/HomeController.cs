using System.Diagnostics;
using ExcelRefinery.Models;
using ExcelRefinery.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authorization;

namespace ExcelRefinery.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IExcelProcessingService _excelProcessingService;

        public HomeController(ILogger<HomeController> logger, IExcelProcessingService excelProcessingService)
        {
            _logger = logger;
            _excelProcessingService = excelProcessingService;
        }

        public IActionResult Index()
        {
            // Check if user is authenticated
            if (User.Identity.IsAuthenticated)
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

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
