using System.ComponentModel.DataAnnotations.Schema;

namespace ImportExcel
{
    [Table("order")]
    public class order : BaseModel, IBaseImportModel
    {
        public string reftype { get; set; }

        public string method { get; set; }
    }
}
