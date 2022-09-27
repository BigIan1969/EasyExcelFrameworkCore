namespace EasyExcelFramework
{
    public class Conditions
    {
        public enum FileFormats { XLSX, XLSB, XLS, CSV }
        public int excelextension(string filename)
        {
            //choose file extension
            switch (Path.GetExtension(filename).ToUpper())
            {
                case "." + nameof(FileFormats.XLS):
                    return (int)FileFormats.XLS;
                case "." + nameof(FileFormats.XLSX):
                    return (int)FileFormats.XLSX;
                case "." + nameof(FileFormats.CSV):
                    return (int)FileFormats.CSV;
                default:
                    return -1;
            }

        }
    }
}
