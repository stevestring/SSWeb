using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace SSWeb.Controllers
{
    public class WorksheetController : Controller
    {
        //
        // GET: /Worksheet/

        public ActionResult Index(int id)
        {

            Models.Worksheet w = new Models.Worksheet();
            w.GetWorksheet(id);
            ViewData.Model = w;
            return View("Worksheets/Details",w);

        }

    }
}
