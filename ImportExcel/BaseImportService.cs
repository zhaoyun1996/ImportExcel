using Dapper;
using Newtonsoft.Json.Linq;
using Npgsql;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ImportExcel
{
    public class BaseImportService<TModel> : IImportService where TModel : IBaseImportModel
    {
        public List<import_column> GetImportColumns(NpgsqlConnection cnn)
        {
            // Tạo câu lệnh SQL Update từ class
            string sqlQuery = $"select * from dbo.import_column where table_id = '{typeof(TModel).Name}' order by sort_order;";
            List<import_column> importColumns = cnn.Query<import_column>(sqlQuery)?.ToList();

            Console.WriteLine("Lấy cấu hình cột.");

            return importColumns;
        }

        public int InsertData<T>(NpgsqlConnection cnn, List<T> data)
        {
            // Lấy các thuộc tính của đối tượng T
            PropertyInfo[] properties = data[0].GetType().GetProperties();

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

            int index = 0;
            int count = 0;

            while (index < data.Count)
            {
                // Xây dựng chuỗi truy vấn SQL
                StringBuilder queryBuilder = new StringBuilder($"INSERT INTO dbo.{data[0].GetType().Name} ({columnNames}) VALUES ");

                int skip = data.Count - index < 500 ? data.Count - index : 500; //data.Count;

                for (int i = 0; i < skip; i++)
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

                    if (i < skip - 1)
                    {
                        queryBuilder.Append(", ");
                    }
                }

                queryBuilder.Append(";");

                string query = queryBuilder.ToString();

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

                index += count;

                Console.WriteLine($"Đã nhập khẩu: {index}/{data.Count}");
            }


            return data.Count;
        }

        public virtual List<BaseModel> GetData(NpgsqlConnection cnn, List<import_column> importColumns)
        {
            return new List<BaseModel>();
        }
    }
}
