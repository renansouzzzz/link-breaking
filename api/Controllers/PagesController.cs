using CallContent.Models;
using CallContent.Service;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace CallContent.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class PagesController : ControllerBase
    {
        private readonly PagesService _page;

        public PagesController(PagesService page)
        {
            _page = page;
        }

        [HttpGet(Name = "GetPages")]
        public async Task<List<SitePage>> GetAsync()
        {
            List<SitePage> pages = await _page.ListPages("4156c839-562e-4702-b7ac-00c97ee6b4a8");

            foreach (var page in pages)
                RunPnPScript(page.Title);

            return pages;
        }

        private void RunPnPScript(string pageName)
        {
            pageName = pageName.Replace(" ", "-") + ".aspx";

            string scriptPath = @"C:\Users\Tnend\Documents\dev-solucoes\k2m\robo-quebra-links\get-content-to-xml1.ps1";

            string pwshPath = @"C:\Program Files\PowerShell\7\pwsh.exe";

            ProcessStartInfo processStartInfo = new ProcessStartInfo
            {
                FileName = pwshPath,
                Arguments = $"-ExecutionPolicy Bypass -File \"{scriptPath}\" -pageName \"{pageName}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (Process process = new Process { StartInfo = processStartInfo })
            {
                process.Start();

                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();

                process.WaitForExit();

                if (!string.IsNullOrEmpty(process.StandardError.ReadToEnd()))
                {
                    Console.WriteLine("Erro ao executar o script PowerShell:");
                    Console.WriteLine(error);
                }
                else
                {
                    Console.WriteLine("Script executado com sucesso:");
                    Console.WriteLine(output);
                }
            }
        }
    }
}
