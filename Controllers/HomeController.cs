using System.Diagnostics;
using ExcelRefinery.Models;
using Microsoft.AspNetCore.Mvc;

namespace ExcelRefinery.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
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

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
