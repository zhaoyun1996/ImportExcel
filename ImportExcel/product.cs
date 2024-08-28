using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportExcel
{
    [Table("product")]
    public class product : BaseModel, IBaseImportModel
    {
        public string sku { get; set; }

        public string image_1 { get; set; }

        public string image_2 { get; set; }

        public string print_area_front { get; set; }

        public string print_area_back { get; set; }

        public string color { get; set; }

        public string size { get; set; }

        public string note { get; set; }
    }
}
