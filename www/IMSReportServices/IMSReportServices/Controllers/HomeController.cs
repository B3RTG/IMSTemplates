using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace IMSReportServices.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.Title = "Lanzar Job";
            ViewBag.ImageResourcePath = "./images/";
            ViewBag.IndexActive = "active";

            return View();
        }

        public ActionResult TaskList()
        {
            ViewBag.Title = "Review Task";
            ViewBag.CurrentView = "TASKLIST";
            ViewBag.TaskListActive = "active";
            ViewBag.ImageResourcePath = "../images/";

            return View();
        }
    }
}
