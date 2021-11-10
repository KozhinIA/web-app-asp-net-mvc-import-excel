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
    public class ImportAuthorData
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public DateTime? Birthday { get; set; }
        public Gender Gender { get; set; }
    }
}