using ExcelService;
using ModelImport;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace WebTest.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            IList<ImportError> errors = new List<ImportError>();
            IList<SampleImport> importList;
            var filePath = RootDirectoryHelper.GetFilePath("./Excels/sampleImport.xlsx");
            var cfgNodeName = "Sample";

            using (var fs = new FileStream(filePath, FileMode.Open))
            {
                importList = new ExcelImportService<SampleImport>().GetParsedPositionImport(fs, errors, cfgNodeName);
            }

            //if (errors.Count > 0)
            //    // handle errors

            return View(importList);
        }
    }
}