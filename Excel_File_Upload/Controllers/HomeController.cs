using Excel_File_Upload.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Excel_File_Upload.Controllers
{
    public class HomeController : Controller
    {
        private readonly FileRepo _fileRepo;
        public HomeController()
        {
            _fileRepo = new FileRepo();
        }
        public ActionResult Index()
        {
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

        [HttpGet]
        public ActionResult FileUpload()
        {
            return View();
        }

        // save the file in db using connection string in data.filerepository using stored procedure
        // create a table in db Person with columns Id, Name, Age, Address etc and save the data in db using stored procedure InsertPersonData
        [HttpPost]
        public ActionResult Upload(HttpPostedFileBase file)
        {
            if (file != null && file.ContentLength > 0)
            {
                try
                {
                    var flag = _fileRepo.SaveFileInDb(file);
                    if (flag)
                    {
                        ViewBag.Message = "File uploaded successfully!!";
                    }
                    else
                    {
                        ViewBag.Message = "File upload failed!!";
                    }
                }
                catch (Exception ex)
                {
                    ViewBag.Message = "File upload failed!!";
                }

            }
            return RedirectToAction("Index");
        }


        [HttpGet]
        public ActionResult GetFileData()
        {
            var data = _fileRepo.GetFileData();
            return View(data);

        }
    }
}