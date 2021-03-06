using System.ComponentModel.DataAnnotations;

namespace WebAppAspNetMvcImportExcel.Models
{
    public enum ImportAuthorRowLogType
    {
        [Display(Name = "Успешно")]
        Success = 1,

        [Display(Name = "Ошибка при парсинге строки")]
        ErrorParsed = 2,
    }
}