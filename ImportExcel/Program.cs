using Microsoft.Extensions.DependencyInjection;
using Newtonsoft.Json.Linq;
using Npgsql;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace ImportExcel
{
    public class Product
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public decimal Price { get; set; }
    }

    public class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Thiết lập mã hóa UTF-8 cho console
                Console.OutputEncoding = Encoding.UTF8;

                // Create a service collection and configure services
                var serviceCollection = new ServiceCollection();
                ConfigureServices(serviceCollection);

                // Build the service provider
                var serviceProvider = serviceCollection.BuildServiceProvider();

                ImportUtil importUtil;

                // Create a scope
                using (var scope = serviceProvider.CreateScope())
                {
                    importUtil = scope.ServiceProvider.GetRequiredService<ImportUtil>();
                }

                Console.WriteLine("Chương trình đang xử lý.");

                string _connectionString = "User ID=postgres; Username=admin; Password=dwmk228ytOLx7Sg54wsUQCVSA8eQmHb3; Host=dpg-cr1lcedsvqrc73dvunag-a.singapore-postgres.render.com; Port=5432; Database=wa_imart;";
                using (var cnn = new NpgsqlConnection(_connectionString))
                {
                    cnn.Open();

                    Console.WriteLine("Kết nối dữ liệu thành công.");

                    IImportService importService = importUtil.GetImportService("product");

                    List<import_column> importColumns = importService.GetImportColumns(cnn);

                    List<BaseModel> data = importService.GetData(cnn, importColumns);

                    importService.InsertData(cnn, data);

                    //Console.WriteLine($"Nhập khẩu thành công {count} dòng.");
                    Console.WriteLine("Nhấn Enter để kết thúc...");
                    Console.ReadLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Đã xảy ra lỗi: " + ex.Message);
            }
        }

        private static void ConfigureServices(IServiceCollection services)
        {
            // Register services
            services.AddScoped<ImportUtil>();
            services.AddScoped<product_import_service>();
            services.AddScoped<order_import_service>();
        }
    }
}
