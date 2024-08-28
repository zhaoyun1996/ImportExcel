using Newtonsoft.Json.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace ImportExcel
{
    public class ExcelImporter
    {
        public List<T> ImportExcelToDataTable<T>(string filePath, List<import_column> importColumns) where T : new ()
        {
            List<T> data = new List<T>();
            DataTable dataTable = new DataTable();
            IWorkbook workbook;

            // Mở file stream
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                // Kiểm tra định dạng file dựa trên phần mở rộng
                if (Path.GetExtension(filePath).Equals(".xls"))
                {
                    workbook = new HSSFWorkbook(fileStream); // Dùng cho .xls
                }
                else if (Path.GetExtension(filePath).Equals(".xlsx"))
                {
                    workbook = new XSSFWorkbook(fileStream); // Dùng cho .xlsx
                }
                else
                {
                    throw new Exception("Định dạng file không được hỗ trợ!");
                }

                ISheet sheet = workbook.GetSheetAt(0); // Lấy sheet đầu tiên
                IRow headerRow = sheet.GetRow(0); // Giả sử hàng đầu tiên là tiêu đề

                int cellCount = headerRow.LastCellNum;
                int rowCount = sheet.LastRowNum;

                // Tạo các cột cho DataTable dựa trên tiêu đề
                for (int i = 0; i < cellCount; i++)
                {
                    ICell cell = headerRow.GetCell(i);
                    if (cell != null)
                    {
                        dataTable.Columns.Add(cell.ToString());
                    }
                    else
                    {
                        dataTable.Columns.Add("Column" + i);
                    }
                }

                // Đọc dữ liệu từ các hàng tiếp theo và thêm vào DataTable
                for (int i = 1; i <= rowCount; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue; // Bỏ qua hàng trống

                    T item = new T();

                    for (int j = 0; j < cellCount; j++)
                    {
                        ICell cell = row.GetCell(j);
                        if (cell != null)
                        {
                            if (cell.CellType == CellType.Formula)
                            {
                                // Đọc giá trị kết quả của công thức
                                var formulaEvaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
                                cell = formulaEvaluator.EvaluateInCell(cell);
                            }

                            var x = importColumns.FirstOrDefault(ic => string.Equals(ic.column_name_excel, dataTable.Columns[j].ToString(), StringComparison.OrdinalIgnoreCase));

                            var pr = item.GetType().GetProperty(x.column_id);
                            if(pr != null)
                            {
                                pr.SetValue(item, GetCellValue(cell));
                            }
                        }
                    }

                    data.Add(item);
                }
            }

            return data;
        }

        private object GetCellValue(ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                        return cell.DateCellValue;
                    else
                        return cell.NumericCellValue;
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.Formula:
                    return cell.ToString(); // Hoặc xử lý công thức nếu cần
                default:
                    return cell.ToString();
            }
        }
    }

}
