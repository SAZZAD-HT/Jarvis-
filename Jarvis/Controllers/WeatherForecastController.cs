using Jarvis.Helper.MailReader;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;  // make sure this is included
using System.Collections.Generic;    // for IEnumerable
using System.Linq;                   // for Select
using System;                        // for Console if still needed (although removed here)
using Microsoft.Office.Interop.Outlook;
namespace Jarvis.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
            "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
        };

        private readonly ILogger<WeatherForecastController> _logger;

        public WeatherForecastController(ILogger<WeatherForecastController> logger)
        {
            _logger = logger;
        }

        [HttpGet(Name = "GetWeatherForecast")]
        public IEnumerable<object> Get()
        {
            var reader = new OutlookDesktopReader();
            var mails = reader.GetUnreadOrderMailsFromRoboFolder();

            var mailSummaries = mails.Select(mail => new
            {
                Subject = mail.Subject,
                ReceivedTime = mail.ReceivedTime,
                SenderEmail = mail.SenderEmail
            }).ToList();

            return mailSummaries;
        }
    }
}
