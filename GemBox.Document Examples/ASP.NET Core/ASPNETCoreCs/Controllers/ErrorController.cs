using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Document.Samples.Core.Controllers
{
    public class ErrorController : Controller
    {
        public string Error()
        {
            return "Error";
        }
    }
}
