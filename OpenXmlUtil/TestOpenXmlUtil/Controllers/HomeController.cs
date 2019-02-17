using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using OpenXmlUtil;

namespace TestOpenXmlUtil.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {

            string model = Server.MapPath("~/Arquivo/Template.docx"); //Utilizamos o modelo para prencher os valores --Aqui passar o tipo dependendo da pratica do parceiro
            string newDocument = Server.MapPath(string.Format("~/Arquivo/Bonus-{0}.docx", "NomeFuncionario")); //Cria a cópia no local Indicado

            if (System.IO.File.Exists(model))
            {
               CoreControl.ProcessDocuments(model, newDocument);
            }
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}