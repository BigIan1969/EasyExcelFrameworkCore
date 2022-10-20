namespace EasyExcelFramework
{
    public class TestStepsLogEntry
    {
        public DateTime Started;
        public DateTime End;
        public bool Outcome;
        public string worksheet;
        public int Rownumber;
        public string Action;
        public string[] parameters;
        public string? screenshot;
        public Exception? Ex;
    }
}
