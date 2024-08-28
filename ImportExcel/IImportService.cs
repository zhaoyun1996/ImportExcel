using Newtonsoft.Json.Linq;
using Npgsql;
using System.Collections.Generic;

namespace ImportExcel
{
    public interface IImportService
    {
        List<import_column> GetImportColumns(NpgsqlConnection cnn);

        List<BaseModel> GetData(NpgsqlConnection cnn, List<import_column> importColumns);

        int InsertData<T>(NpgsqlConnection cnn, List<T> data);
    }
}
