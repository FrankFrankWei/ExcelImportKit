using ExcelService;
using ModelImport;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelImportKit
{
    class Program
    {
        static void Main(string[] args)
        {
            var fs = new FileStream(@"F:\\sampleImport.xlsx", FileMode.Open);
            IList<ImportError> errors = new List<ImportError>();
            var importList = new SampleImportService<SampleImport>().GetParsedPositionImport(fs, errors);

            if (errors.Count > 0)
                errors.ToList().ForEach(item => Console.WriteLine($"{item.Line} - {item.ErrorMsg}"));

            importList.Where(m => !m.IsError).ToList().ForEach(item => Console.WriteLine($"{item.Age} - {item.Name} - {item.Height} - " +
                $"{item.GenderName} - {item.Birthday} - {item.Money} - {item.StateName}"));


            Console.Read();
        }
    }
}
