using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportExcel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Thiết lập mã hóa UTF-8 cho console
                Console.OutputEncoding = Encoding.UTF8;

                Console.WriteLine("Chương trình đang xử lý.");

                ExcelImporter importer = new ExcelImporter();
                string filePath = @"C:\Users\vvkiet\Desktop\Mau_ban_hang_da_tien_te.xls";

                DataTable dataTable = importer.ImportExcelToDataTable(filePath);

                // Hiển thị dữ liệu
                foreach (DataRow row in dataTable.Rows)
                {
                    foreach (var item in row.ItemArray)
                    {
                        Console.Write(item + "\t");
                    }
                    Console.WriteLine();
                }

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
