namespace EasyExcelFramework
{
    public class TestLog
    {
        public DateTime Started;
        public DateTime End;
        public bool Outcome;
        public string Test;
        public dynamic[] parameters;
        public List<TestStepsLogEntry> StepHistory;

    }
}
