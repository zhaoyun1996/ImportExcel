using Dapper;
using Newtonsoft.Json.Linq;
using Npgsql;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.Remoting.Contexts;
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
                using (var cnn = new NpgsqlConnection(_connectionString))
                {
                    cnn.Open();

                    Console.WriteLine("Kết nối dữ liệu thành công.");

                    List<import_column> importColumns = null;

                    // Tạo câu lệnh SQL Update từ class
                    string sqlQuery = "select * from dbo.import_column;";
                    importColumns = cnn.Query<import_column>(sqlQuery)?.ToList();

                    Console.WriteLine("Lấy cấu hình cột.");

                    ExcelImporter importer = new ExcelImporter();
                    string filePath = @"C:\Users\vuvan\Desktop\AECoder.xlsx";

                    List<product> data = importer.ImportExcelToDataTable<product>(filePath, importColumns);

                    Console.WriteLine("Đọc xong dữ liệu từ file excel.");

                    //product item = new product();
                    //var getProperties = item.GetType().GetProperties();
                    //List<string> pro = new List<string>();

                    // Lấy các thuộc tính của đối tượng T
                    var properties = typeof(product).GetProperties();

                    // Xây dựng chuỗi các cột
                    StringBuilder columnNames = new StringBuilder();
                    for (int i = 0; i < properties.Length; i++)
                    {
                        columnNames.Append(properties[i].Name);

                        if (i < properties.Length - 1)
                        {
                            columnNames.Append(", ");
                        }
                    }

                    // Xây dựng chuỗi truy vấn SQL
                    StringBuilder queryBuilder = new StringBuilder($"INSERT INTO dbo.product ({columnNames}) VALUES ");

                    int maxCount = 500; //data.Count;

                    for (int i = 0; i < maxCount; i++)
                    {
                        queryBuilder.Append("(");
                        for (int j = 0; j < properties.Length; j++)
                        {
                            queryBuilder.Append($"@{properties[j].Name}{i}");

                            if (j < properties.Length - 1)
                            {
                                queryBuilder.Append(", ");
                            }
                        }
                        queryBuilder.Append(")");

                        if (i < maxCount - 1)
                        {
                            queryBuilder.Append(", ");
                        }
                    }

                    queryBuilder.Append(";");

                    string query = queryBuilder.ToString();

                    var count = 0;

                    using (var command = new NpgsqlCommand(query, cnn))
                    {
                        // Thêm tham số vào lệnh
                        for (int i = 0; i < data.Count; i++)
                        {
                            foreach (var property in properties)
                            {
                                var value = property.GetValue(data[i]);
                                command.Parameters.AddWithValue($"@{property.Name}{i}", value ?? DBNull.Value);
                            }
                        }

                        count = command.ExecuteNonQuery();
                    }

                    Console.WriteLine($"Nhập khẩu thành công {count} dòng.");
                    Console.WriteLine("Nhấn Enter để kết thúc...");
                    Console.ReadLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Đã xảy ra lỗi: " + ex.Message);
            }
        }
    }
}
