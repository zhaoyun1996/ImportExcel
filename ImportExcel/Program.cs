using Dapper;
using Newtonsoft.Json.Linq;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace ImportExcel
{
    public class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Thiết lập mã hóa UTF-8 cho console
                Console.OutputEncoding = Encoding.UTF8;

                Console.WriteLine("Chương trình đang xử lý.");

                string _connectionString = "User ID=postgres; Username=admin; Password=dwmk228ytOLx7Sg54wsUQCVSA8eQmHb3; Host=dpg-cr1lcedsvqrc73dvunag-a.singapore-postgres.render.com; Port=5432; Database=wa_imart;";
                NpgsqlConnection connection = new NpgsqlConnection(_connectionString);
                connection.Open();

                Console.WriteLine("Kết nối dữ liệu thành công.");

                List<import_column> importColumns = null;

                using (IDbConnection db = new NpgsqlConnection(_connectionString))
                {
                    // Tạo câu lệnh SQL Update từ class
                    string sqlQuery = "select * from dbo.import_column;";

                    importColumns = db.Query<import_column>(sqlQuery)?.ToList();
                }

                Console.WriteLine("Lấy cấu hình cột import_column.");

                ExcelImporter importer = new ExcelImporter();
                string filePath = @"C:\Users\vvkiet\Desktop\Mau_ban_hang_da_tien_te.xls";

                List<order> dataTable = importer.ImportExcelToDataTable(filePath, importColumns);

                // Hiển thị dữ liệu
                //foreach (DataRow row in dataTable.Rows)
                //{
                //    foreach (var item in row.ItemArray)
                //    {
                //        Console.Write(item + "\t");
                //    }
                //    Console.WriteLine();
                //}

                Console.WriteLine("Chương trình đã hoàn thành.");
                Console.WriteLine("Nhấn Enter để kết thúc...");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Đã xảy ra lỗi: " + ex.Message);
            }
        }
    }
}
