using CallContent.Models;
using CallContent.Service;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;

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

        [HttpGet("/get-links")]
        public async Task<List<LinkInfo>> GetLinksAsync()
        {
            List<LinkInfo> links = new List<LinkInfo>();

            foreach (var page in _page.GetSharepointPageId().Result)
            {
                var contentLinks = await _page.GetContent(page);
                links.AddRange(contentLinks);
            }

            await _page.SaveLinksToExcelAsync(links, @"C:\dev\k2m\links-quebrados.xlsx");

            return links;
        }
    }
}
