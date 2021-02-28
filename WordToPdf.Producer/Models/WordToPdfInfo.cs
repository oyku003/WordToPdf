using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WordToPdf.Producer.Models
{
    public class WordToPdfInfo
    {
        public string Email { get; set; }
        public IFormFile WordFile { get; set; }
    }
}
