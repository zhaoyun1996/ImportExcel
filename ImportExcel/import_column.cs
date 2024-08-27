using System;

namespace ImportExcel
{
    public class import_column
    {
        public Guid import_column_id { get; set; }

        public int sort_order { get; set; }

        public bool hiden_system { get; set; }

        public string table_id { get; set; }

        public string column_id { get; set; }

        public string column_name { get; set; }

        public string column_name_excel { get; set; }

        public string column_name_excel_mapping { get; set; }

        public string data_type { get; set; }

        public string validate { get; set; }

        public bool shown { get; set; }
    }
}
