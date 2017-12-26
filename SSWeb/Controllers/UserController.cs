using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using System.Web.Security;
namespace SSWeb.Controllers
{
    [Authorize]
    public class UserController: Controller
    {

        public ActionResult Index()
        {
           //var user = Membership.GetUser();
            Models.User u = new Models.User();
            u.UserID = User.Identity.Name;
            u.GetOwnedWorksheets();
            ViewData.Model = u;
            return View();

        }



    }
}