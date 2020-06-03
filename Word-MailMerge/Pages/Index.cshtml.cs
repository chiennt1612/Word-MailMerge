using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Logging;
using Utils;

namespace Word_MailMerge.Pages
{
    public class IndexModel : PageModel
    {
        private readonly ILogger<IndexModel> _logger;

        public IndexModel(ILogger<IndexModel> logger)
        {
            _logger = logger;
        }

        public void OnGet()
        {
            _logger.Log(LogLevel.Information, "Starting Object");
            string SourcePath = Directory.GetCurrentDirectory() + "\\wwwroot\\Template\\" + "phieuluong.docx";
            string DestinationPath = Directory.GetCurrentDirectory() + "\\wwwroot\\Template\\" + "result.docx";
            string data = LogHelper.FileReadAllText(Directory.GetCurrentDirectory() + "\\" + "Data.json");
            var a = new WordProcessing(SourcePath, DestinationPath, data);

            _logger.Log(LogLevel.Information, "End Object");
        }
    }
}
