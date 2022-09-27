using ExtratorCTRC.Business;
using ExtratorCTRC.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace ExtratorCTRC.Controllers
{
  public class HomeController : Controller
  {
    private readonly ILeituraArquivosBLL _leituraArquivoBLL;

    public HomeController(ILeituraArquivosBLL leituraArquivosBLL)
    {
      _leituraArquivoBLL = leituraArquivosBLL;
    }

    public IActionResult Index(ArquivoModel arq)
    {
      DateTime value = new DateTime(DateTime.Now.Year, 10, 3);

      try
      {
        var msgSucessoCriarPlanilha = string.Empty;

        if (DateTime.Now <= value)
        {
          if (arq.Arquivos is not null)
            msgSucessoCriarPlanilha = _leituraArquivoBLL.RetornaCamposHtml(arq.Arquivos);
        }

        if (!string.IsNullOrEmpty(msgSucessoCriarPlanilha))
          TempData["ShowAlert"] = msgSucessoCriarPlanilha;
      }
      catch (Exception ex)
      {
        TempData["AlertMessage"] = ex.Message;
      }

      return View();
    }


    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
      return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }
  }
}