using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using RabbitMQ.Client;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordToPdf.Producer.Models;

namespace WordToPdf.Producer.Controllers
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
            return View();
        }

        public IActionResult WordToPdfPage()
            => View();

        [HttpPost]
        public IActionResult WordToPdfPage(WordToPdfInfo wordToPdf)
        {
            var factory = new ConnectionFactory { HostName = "localhost" };
            using (var connection = factory.CreateConnection())
            {
                using (var channel = connection.CreateModel())
                {
                    channel.ExchangeDeclare("convert-exchange", ExchangeType.Direct, true, false, null);
                    //exclusive:false-> birden fazla bağlantı bu kuyrugu kullanabilsin dedik
                    channel.QueueDeclare(queue: "File", durable: true, exclusive: false, autoDelete: false, arguments: null);

                    channel.QueueBind("File", "convert-exchange", "WordToPdf");

                    MessageWordToPdf messageWordToPdf = new MessageWordToPdf();

                    using (MemoryStream ms = new MemoryStream())
                    {
                        wordToPdf.WordFile.CopyTo(ms);//hafızaya yazdık
                        messageWordToPdf.WordByte = ms.ToArray();
                    }

                    messageWordToPdf.Email = wordToPdf.Email;
                    messageWordToPdf.FileName = Path.GetFileNameWithoutExtension(wordToPdf.WordFile.FileName);

                    string serialize = JsonConvert.SerializeObject(messageWordToPdf);
                    byte[] byteMessage = Encoding.UTF8.GetBytes(serialize);
                    var properties = channel.CreateBasicProperties();
                    properties.Persistent = true;//rabbitmq instance'ı restart olsa dahi mesajım kaybolmayacak
                    channel.BasicPublish("convert-exchange", "WordToPdf", properties, byteMessage);

                    ViewBag.Result = "Word dosyanız pdf dosyasına dönüştürüldükten sonra size email olarak gönderilecektir.";
                    return View();

                }

            }
        }

        public IActionResult Privacy()
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
