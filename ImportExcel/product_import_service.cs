using Newtonsoft.Json.Linq;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace ImportExcel
{
    public class product_import_service : BaseImportService<product>, IImportService
    {
        public override List<BaseModel> GetData(NpgsqlConnection cnn, List<import_column> importColumns)
        {
            ExcelImporter importer = new ExcelImporter();

            string filePath = @"C:\Users\vvkiet\Desktop\AECoder.xlsx";

            List<BaseModel> data = new List<BaseModel>();

            List<product> products = importer.ImportExcelToDataTable<product>(filePath, importColumns);

            data.AddRange(products);

            Console.WriteLine("Đọc xong dữ liệu từ file excel.");

            return data;
        }
    }
}
