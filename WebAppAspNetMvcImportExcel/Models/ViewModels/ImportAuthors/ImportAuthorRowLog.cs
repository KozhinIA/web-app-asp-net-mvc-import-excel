using Common.Attributes;
using Common.Extensions;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Web;
using System.Web.Mvc;

namespace WebAppAspNetMvcImportExcel.Models
{
    public class ImportAuthorRowLog
    {
        public int Id { get; set; }
        public string Message { get; set; }
        public ImportAuthorRowLogType Type { get; set; }
    }
}