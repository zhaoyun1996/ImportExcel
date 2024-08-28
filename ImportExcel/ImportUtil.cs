namespace ImportExcel
{
    public class ImportUtil
    {
        private readonly product_import_service _product;
        private readonly order_import_service _order;

        public ImportUtil(product_import_service product, order_import_service order)
        {
            _product = product;
            _order = order;
        }

        public IImportService GetImportService(string tableName)
        {
            IImportService service = null;
            switch (tableName)
            {
                case "order":
                    service = _order;
                    break;
                case "product":
                    service = _product;
                    break;
                default:
                    break;
            }

            return service;
        }
    }
}
