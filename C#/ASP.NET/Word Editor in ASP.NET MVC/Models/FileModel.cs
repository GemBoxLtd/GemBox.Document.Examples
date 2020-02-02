using System.Web.Mvc;

namespace AspNetWordEditor.Models
{
    public sealed class FileModel
    {
        [AllowHtml]
        public string Content { get; set; }

        public string Extension { get; set; }
    }
}