using AiderHubAtual.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Web;
using System.Threading.Tasks;

namespace AiderHubAtual.Controllers
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

        public IActionResult Privacy()
        {
            return View();
        }

        //[HttpPost]
        //public ActionResult ExecutaMacro()
        //{
        //    string caminho = "C:\\Users\\PC\\source\\repos\\AiderHubAtual\\AiderHubAtual\\Relatorio\\MacroCertificado.xlsm";

        //    Application xlApp = new Application();

        //    if (xlApp == null)
        //    {
        //        ViewBag.Mensagem = "Erro ao executar a macro: aplicativo Excel não encontrado.";
        //        return View("Relatorio");
        //    }

        //    Workbook xlWorkbook = xlApp.Workbooks.Open(caminho, ReadOnly: false);

        //    try
        //    {
        //        xlApp.Visible = false;
        //        xlApp.Run("GerarCertificado");
        //    }
        //    catch (System.Exception)
        //    {
        //        ViewBag.Mensagem = "Erro ao executar a macro.";
        //        return View("Relatorio");
        //    }

        //    xlWorkbook.Close(false);
        //    xlApp.Application.Quit();
        //    xlApp.Quit();


        //    ViewBag.Mensagem = "Arquivo gerado com sucesso!";
        //    return View("Relatorio");
        //}

        public ActionResult Relatorio()
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
