using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
namespace SSWeb.Controllers
{
    public class WorksheetsController: Controller
    {
        [Authorize]
        public ActionResult Upload()
        {
            return View();
        }
        public ActionResult UploadDemo()
        {
            return View();
        }

        // This action handles the form POST and the upload
        [HttpPost]
        [Authorize]
        public ActionResult Upload(HttpPostedFileBase file)
        {
            // Verify that the user selected a file
            if (file != null && file.ContentLength > 0 && file.ContentLength < 1048576)
            {
                if ( Path.GetExtension(file.FileName) == ".xlsx")//make sure excel
                {
                    // extract only the fielname
                    var fileName = Path.GetFileName(file.FileName);
                    //// store the file inside ~/App_Data/uploads folder
                    var path = Path.Combine(HttpContext.Server.MapPath("~/App_Data"), fileName);

                    file.SaveAs(path);

                    Models.Worksheet w = new Models.Worksheet();
                    w.Owner = User.Identity.Name;
                    int i = w.PutWorksheetStub(Path.GetFileName(path));
                    w.WorksheetId = i;
                    w.Name = Path.GetFileNameWithoutExtension(path);


                    SpreadsheetUploader.SpreadsheetUploader su = new SpreadsheetUploader.SpreadsheetUploader();

                    su.LockWorksheet(w.WorksheetId, true);
                    su.LoadSpreadsheet(i,path);
                    su.LockWorksheet(w.WorksheetId, false);
                    su.SetWorkSheetComplete(w.WorksheetId);

                    return RedirectToAction("EditNew", new { id = i });

                }
                else
                {
                    ViewData.Add("ErrorMessage", "This file does not appear to be a valid Excel file.");
                }
            }
            else
            {
                ViewData.Add("ErrorMessage", "This file exceeds 1MB.");
            }
            // redirect back to the index action to show the form once again
            //return RedirectToAction("Index");
            return View();
        }

        // This action handles the form POST and the upload
        [HttpPost]
        public ActionResult UploadDemo(HttpPostedFileBase file)
        {
            // Verify that the user selected a file
            if (file != null && file.ContentLength > 0)
            {
                if (Path.GetExtension(file.FileName) == ".xls"
                    || Path.GetExtension(file.FileName) == ".xlsx")//make sure excel
                {
                    // extract only the fielname
                    var fileName = Path.GetFileName(file.FileName);
                    // store the file inside ~/App_Data/uploads folder
                    var path = Path.Combine(HttpContext.Server.MapPath("~/App_Data"), fileName);

                    file.SaveAs(path);

                    Models.Worksheet w = new Models.Worksheet();
                    w.Owner = "Demo";
                    int i = w.PutWorksheetStub(Path.GetFileName(path));
                    w.WorksheetId = i;
                    w.Name = Path.GetFileNameWithoutExtension(path);
                    
                    return RedirectToAction("EditDemo", new { id = i });

                }
                else
                {
                    ViewData.Add("ErrorMessage","This file does not appear to be a valid Excel file.");
                }
            }
            else
            {
                ViewData.Add("ErrorMessage", "This file exceeds 1MB.");
            }
            // redirect back to the index action to show the form once again
            //return RedirectToAction("Index");
            return View();
        }


        public ActionResult Worksheets(int id)
        {
            Models.Worksheet w = new Models.Worksheet();
            w.GetWorksheet(id);
            ViewData.Model = w;
            return View("Details", w);
            
        }

        public ActionResult Index()
        {

            Models.Worksheets w = new Models.Worksheets();

            w.GetWorkSheets();
            ViewData.Model = w;
            //return View("Details", w);
            return View();

        }



        [Authorize]
        public ActionResult Edit(int id)
        {
            Models.Worksheet w = new Models.Worksheet();
            w.GetWorksheet(id);

            if (w.Owner.ToUpper() != User.Identity.Name.ToUpper())
            {
                return RedirectToAction("Unauthorized");
            }

            ViewData.Model = w;
            return View();
        }

        [Authorize]
        public ActionResult EditNew(int id)
        {
            Models.Worksheet w = new Models.Worksheet();
            w.GetWorksheet(id);

            if (w.Owner.ToUpper() != User.Identity.Name.ToUpper())
            {
                return RedirectToAction("Unauthorized");
            }

            ViewData.Model = w;
            return View();
        }

        [Authorize]
        public ActionResult Delete(int id)
        {
            Models.Worksheet w = new Models.Worksheet();
            w.GetWorksheet(id);
            if (w.Owner.ToUpper() != User.Identity.Name.ToUpper())
            {
                return RedirectToAction("Unauthorized");
            }

            w.DeleteWorksheet(id);
            return Redirect(Request.UrlReferrer.ToString());//This should be User Page
        }

        public ActionResult Unauthorized()
        {
            return View();
        }


        public ActionResult EditDemo(int id)
        {
            Models.Worksheet w = new Models.Worksheet();
            w.GetWorksheet(id);
            if (w.Owner.ToUpper() != "DEMO")
            {
                return RedirectToAction("Unauthorized");
            }

            ViewData.Model = w;
            return View();
        }
        /// <summary>
        /// This returns the actual worksheet with cells
        /// Partial View for async wait message while loading
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public ActionResult Worksheet(int id)
        {

            Models.Worksheet w = new Models.Worksheet();
            w.GetWorksheet(id);
            int complete = w.Complete;
            int numTries = 0;
            while (complete == 0 && numTries <20)
            {
                System.Threading.Thread.Sleep(500);
                w.GetWorksheet(id);
                complete = w.Complete;
                numTries++;
            }
            w.GetWorksheetCells(id);//Cells exist 
            ViewData.Model = w;
            return PartialView(w);
        }


        public ActionResult Details(int id)
        {
            Models.Worksheet w = new Models.Worksheet();
            w.GetWorksheet(id);

            if (!w.Private || User.Identity.Name.ToUpper() == w.Owner.ToUpper()) //public or match
            {
                ViewData.Model = w;
                return View();
            }
            else if (w.Private && !User.Identity.IsAuthenticated) //private, not logged in
            {
                return RedirectToAction("LogOn", "Account", new { returnUrl = "/Worksheet/" + id });
            }
            else //Logged on, wrong user
            {
                return View("UnauthorizedDetail");                
            }
        }

        public ActionResult DetailsDemo(int id)
        {
            Models.Worksheet w = new Models.Worksheet();
            w.GetWorksheet(id);
            if (w.Owner.ToUpper() != "DEMO")
            {
                return RedirectToAction("Unauthorized");
            }
            ViewData.Model = w;
            return View();
        }


        public ActionResult DetailsLoading(int id)
        {
            Models.Worksheet w = new Models.Worksheet();
            w.GetWorksheet(id);
            ViewData.Model = w;

            return View();
        }

        [Authorize]
        [HttpPost]
        public ActionResult EditNew(Models.Worksheet w)
        {

            if (ModelState.IsValid)
            {
                w.UpdateWorksheet(w.WorksheetId, w.Name, w.Description, User.Identity.Name, w.Private);
                return RedirectToAction("Details", new { id = w.WorksheetId });
            }

            return View(w);

        }

        [Authorize]
        [HttpPost]
        public ActionResult Edit(Models.Worksheet w)
        {

            if (ModelState.IsValid)
            {
                w.UpdateWorksheet(w.WorksheetId, w.Name, w.Description, User.Identity.Name, w.Private);
                return RedirectToAction("Index", "User");
            }

            return View(w);
        }

        [HttpPost]
        public ActionResult EditDemo(Models.Worksheet w)
        {

            if (ModelState.IsValid)
            {

                w.UpdateWorksheet(w.WorksheetId, w.Name, w.Description, "Demo", false);
                return RedirectToAction("DetailsDemo", new { id = w.WorksheetId });
            }  

            return View(w);
        }

    }
}