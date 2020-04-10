using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace DocumentCore.Controllers
{
    public class ErrorController : Controller
    {
        public string Error()
        {
            return "Error";
        }
    }
}
