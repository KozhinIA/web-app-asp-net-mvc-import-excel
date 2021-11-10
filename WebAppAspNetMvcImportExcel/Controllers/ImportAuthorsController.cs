using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web.Mvc;
using WebAppAspNetMvcImportExcel.Models;

namespace WebAppAspNetMvcImportExcel.Controllers
{
    public class ImportAuthorsController : Controller
    {
        [HttpGet]
        public ActionResult Index()
        {
            var model = new ImportAuthorViewModel();
            return View(model);
        }

        [HttpPost]
        public ActionResult Import(ImportAuthorViewModel model)
        {
            if (!ModelState.IsValid)
                return View(model);

            var log = ProceedImport(model);

            return View("Log", log);
        }

        public ActionResult GetExample()
        {
            return File("~/Content/Files/ImportAuthorsExample.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ImportAuthorsExample.xlsx");
        }

        private ImportAuthorLog ProceedImport(ImportAuthorViewModel model)
        {
            var startTime = DateTime.Now;

            var workBook = new XLWorkbook(model.FileToImport.InputStream);
            var workSheet = workBook.Worksheet(1);
            var rows = workSheet.RowsUsed().Skip(1).ToList();

            var logs = new List<ImportAuthorRowLog>();
            var data = ParseRows(rows, logs);
            ApplyImported(data);

            var successCount = data.Count();
            var failedCount = rows.Count() - successCount;
            var finishTime = DateTime.Now;

            var result = new ImportAuthorLog()
            {
                StartImport = startTime,
                EndImport = finishTime,
                SuccessCount = successCount,
                FailedCount = failedCount,
                Logs = logs
            };

            return result;
        }

        private List<ImportAuthorData> ParseRows(IEnumerable<IXLRow> rows, List<ImportAuthorRowLog> logs)
        {
            var result = new List<ImportAuthorData>();
            int index = 1;
            foreach(var row in rows)
            {
                try
                {
                    var data = new ImportAuthorData()
                    {
                        FirstName = ConvertToString(row.Cell("A").GetValue<string>().Trim()),
                        LastName = ConvertToString(row.Cell("B").GetValue<string>().Trim()),
                        Birthday = ConvertToDateTime(row.Cell("C").GetValue<string>().Trim()),
                        Gender = ConvertToGender(row.Cell("D").GetValue<string>().Trim()),
                    };

                    result.Add(data);
                    logs.Add(new ImportAuthorRowLog()
                    {
                        Id = index,
                        Message = $"ОК",
                        Type = ImportAuthorRowLogType.Success 
                    }); ;

                }
                catch (Exception ex)
                {
                    logs.Add(new ImportAuthorRowLog()
                    {
                        Id = index,
                        Message = $"Error: {ex.GetBaseException().Message}",
                        Type = ImportAuthorRowLogType.ErrorParsed
                    }); ;
                }

                index++;
            }


            return result;
        }

        private void ApplyImported(List<ImportAuthorData> data)
        {
            var db = new LibraryContext();
                       
            foreach (var value in data)
            {
                var model = new Author()
                {
                    FirestName = value.FirstName,
                    LastName = value.LastName,
                    Birthday = value.Birthday,
                    Gender = value.Gender                    
                };

                db.Authors.Add(model);
                db.SaveChanges();
            }
        }

        private string ConvertToString(string value)
        {
            if (string.IsNullOrEmpty(value))
                throw new Exception("Значение не определено");

            var reuslt = HandleInjection(value);

            return reuslt;
        }
        private string HandleInjection(string value)
        {
            var badSymbols = new Regex(@"^[+=@-].*");
            return Regex.IsMatch(value, badSymbols.ToString()) ? string.Empty : value;
        }

        private DateTime? ConvertToDateTime(string value)
        {
            if (string.IsNullOrEmpty(value))
                return null;

            DateTime result = default;

            if (DateTime.TryParse(value, out DateTime temp))
                result = temp;

            if (result == default)
                return null;

            return result;
        }

        private Gender ConvertToGender(string value)
        {
            if (string.IsNullOrEmpty(value))
                throw new Exception("Значение не определено");

            Gender result = default;

            if (Enum.TryParse(value, out Gender temp))
                result = temp;

            if (result == default)
                throw new Exception("Значение пола не распознано");

            return result;
        }
    }
}