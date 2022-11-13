using ExcelDataReader;
using System.Collections;
using System.Data;


namespace EasyExcelFramework
{
    public class EasyExcelF
    {
        public EasyExcelF parent;
        //Dictionary of Worksheets
        private Dictionary<string, DataTable>? worksheets;
        public Dictionary<string, DataTable>? Worksheets { get => worksheets; }

        public InterpreterClass Interpreter;

        //worksheet property
        private string worksheet;
        public string Worksheet { get => worksheet; }

        //Dictionary of variables
        public Dictionary<string, dynamic>? Locals;

        //Dictionary of Globals
        public Dictionary<string, dynamic>? Globals;

        public IDictionary? Environ;
        //First worksheet where the framework begins executing

        //RegisteredActions
        private Dictionary<string, Func<EasyExcelF, string[], bool>> registeredactions;

        private string firstworksheet;

        //Current indent the framework is operating at
        private int currentindent;
        public int CurrentIndent { get => currentindent; }

        private string[] currentdatarow;

        public string[] Currentdatarow { get => currentdatarow; }
        public int Currentrownumber { get => currentrownumber; }

        private int currentrownumber;

        public string[] passedparams;

        public List<TestStepsLogEntry> StepHistory;

        public List<TestLog> TestHistory;

        public bool loopactive;
        public bool breakoutofloop;

        public Func<string, string> Screenshot { get => screenshot; set => screenshot = value; }
        private Func<string, string> screenshot;
        private Func<EasyExcelF, string> report;
        private string defaultpath;
        public void finish()
        {
            Console.WriteLine("Passed: " + TestHistory.Count(n => n.Outcome));
            Console.WriteLine("Failed: " + TestHistory.Count(n => !n.Outcome));
            Console.WriteLine();
            bool failed = false;
            foreach (var t in TestHistory)
            {
                Console.WriteLine(string.Format("Test {0,-20} {1,4} {2} {3} {4}",
                    t.Test,
                    t.Outcome ? "Pass" : "Fail",
                    t.Started,
                    t.End - t.Started,
                    string.Join(", ", t.parameters)
                    ));
                if (!t.Outcome)
                    failed = true;
            }
            if (failed)
            {
                Console.WriteLine();
                Console.WriteLine("Detail:");
                Console.WriteLine();
                foreach (var t in TestHistory)
                {
                    if (!t.Outcome)
                    {
                        Console.WriteLine(string.Format("Test: {0,-20} {1,4} {2} {3} {4}",
                            t.Test,
                            t.Outcome ? "Pass" : "Fail",
                            t.Started,
                            t.End - t.Started,
                            string.Join(", ", t.parameters)
                            ));
                        foreach (var s in t.StepHistory)
                        {
                            if (s.Outcome)
                                Console.ForegroundColor = ConsoleColor.Green;
                            else
                                Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine(string.Format("\tStep: {0} {1} {2,-20} {3} {4}",
                                s.Started,
                                s.End - s.Started,
                                s.Action,
                                string.Join(", ", s.parameters),
                                s.Outcome ? "Pass" : "Fail"
                                ));
                            Console.ResetColor();
                            if (!s.Outcome)
                                Console.WriteLine("\t\t"+s.Ex);
                        }
                    }
                }
            }
        }
        public EasyExcelF(string filename = "default.xlsx", string defaultpath = null)
        {
            //Add core
            EECore eec = new EECore(this);
            EELogic eel = new EELogic(this);
            EEAssert eea = new EEAssert(this);

            //Instanciate Interpreter
            Interpreter = new InterpreterClass();

            //Instanciate worksheets
            worksheets = new Dictionary<string, DataTable>(StringComparer.OrdinalIgnoreCase);

            //get environment variables
            Environ = Environment.GetEnvironmentVariables();

            //set indent level
            currentindent = 0;
            //Instanciate variables
            Locals = new Dictionary<string, dynamic>(StringComparer.OrdinalIgnoreCase);

            if (Globals == null)
                Globals = new Dictionary<string, dynamic>(StringComparer.OrdinalIgnoreCase);

            Populateworksheets(filename);
            if (string.IsNullOrEmpty(firstworksheet))
                throw (new Exception("First worksheet not found: " + filename));
            TestHistory = new List<TestLog>();
            if (defaultpath == null)
            {
                this.defaultpath = Directory.GetCurrentDirectory();
            }
            else
            {
                this.defaultpath = defaultpath;
            }
            Console.WriteLine("Easy Excel Framework");
            Console.WriteLine("====================");
            Console.WriteLine("\nTest Initialised: " + DateTime.Now);
            Console.WriteLine(string.Format("File: {0}\n",filename));
        }
        public EasyExcelF()
        {

            //Add core
            EECore eec = new EECore(this);
            EELogic eel = new EELogic(this);
            EEAssert eea = new EEAssert(this);

            //Instanciate Interpreter
            Interpreter = new InterpreterClass();

            //Instanciate worksheets
            worksheets = new Dictionary<string, DataTable>(StringComparer.OrdinalIgnoreCase);
            currentindent = 0;

            //get environment variables
            Environ = Environment.GetEnvironmentVariables();

            //Instanciate variables
            Locals = new Dictionary<string, dynamic>(StringComparer.OrdinalIgnoreCase);
            if (Globals == null)
                Globals = new Dictionary<string, dynamic>(StringComparer.OrdinalIgnoreCase);

            Populateworksheets("Default.xlsx");
            if (string.IsNullOrEmpty(firstworksheet))
                throw (new Exception("First worksheet not found: Default.xslx"));
            TestHistory = new List<TestLog>();
            this.defaultpath = Directory.GetCurrentDirectory();
            Console.WriteLine("Easy Excel Framework");
            Console.WriteLine("====================");
            Console.WriteLine("\nTest Initialised: " + DateTime.Now);
            Console.WriteLine("File: Default.xlsx\n");

        }
        public EasyExcelF(EasyExcelF parent)
        {
            this.parent = parent;
            registeredactions = parent.registeredactions;

            //Instanciate Interpreter
            Interpreter = new InterpreterClass();

            //Instanciate worksheets
            worksheets = parent.worksheets;
            currentindent = 0;
            Globals = parent.Globals;

            Environ = parent.Environ;

            //Instanciate variables
            Locals = new Dictionary<string, dynamic>(StringComparer.OrdinalIgnoreCase);
            TestHistory = new List<TestLog>();
            this.defaultpath = Directory.GetCurrentDirectory();
        }
        public void Execute(string worksheet = null, string[] passedparameters = null)
        {

            //handle null parameter
            if (worksheet is null)
            {
                worksheet = firstworksheet;
            }
            if (string.IsNullOrEmpty(worksheet))
                throw new InvalidDataException("Cannot call blank worksheet");

            if (passedparameters is null)
            {
                passedparameters = new string[1];
            }
            passedparams = passedparameters;
            //handle null worksheets
            if (worksheets == null)
            {
                throw new ArgumentNullException(nameof(worksheets));
            }
            if (Globals == null)
                Globals = new Dictionary<string, dynamic>(StringComparer.OrdinalIgnoreCase);

            this.worksheet = worksheet;

            this.currentrownumber = 0;

            //loop through specified worksheet
            foreach (DataRow execrow in worksheets[worksheet].Rows)
            {
                TestStepsLogEntry steplog = new TestStepsLogEntry();
                this.currentrownumber++;
                if (execrow[currentindent] == null)
                    throw new ArgumentNullException(nameof(execrow));

                //Convert ItemArray into string array
                this.currentdatarow = Array.ConvertAll(execrow.ItemArray, x => x.ToString());

                int ind = this.currentindent + 1;
                string[] parms = this.currentdatarow[ind..];
                string[] processedparams = new string[parms.Length];
                
                steplog.Started = DateTime.Now;
                steplog.Outcome = false;
                steplog.Action = execrow[currentindent].ToString();
                steplog.parameters = parms;
                steplog.worksheet = worksheet;
                steplog.Rownumber = Currentrownumber;
                string cleanaction;
                bool ignoreerrors = false;
                if (steplog.Action[0].ToString()=="*")
                {
                    ignoreerrors = true;
                    cleanaction = steplog.Action[1..];
                    this.currentdatarow[currentindent] = this.currentdatarow[currentindent][1..];
                }
                else
                {
                    cleanaction = steplog.Action;
                }
                if (StepHistory != null)
                    StepHistory.Add(steplog);
                for (int i = 0; i < parms.Length; i++)
                {
                    if (parms[i].StartsWith("="))
                    {
                        try
                        {
                            processedparams[i] = Interpreter.EvalToString(this, parms[i][1..], passedparameters);
                        }
                        catch
                        {
                            processedparams[i] = parms[i];
                        }
                    }
                    else
                    {
                        processedparams[i] = parms[i];
                    }
                }
                try
                {
                    if (registeredactions.ContainsKey(cleanaction))
                    {
                        bool result = registeredactions[cleanaction](this, processedparams);
                    }
                    else
                    {
                        //If it's a worksheet
                        if (worksheets.ContainsKey(cleanaction) ||
                                                   cleanaction.ToUpper() == "CALL" ||
                                                   cleanaction.ToUpper() == "TEST")
                        {
                            calltestcase(this.currentdatarow);
                        }
                        else
                        {
                            if (!ignoreerrors)
                                throw new InvalidOperationException("Unrecognised Action: " + execrow[currentindent].ToString());
                        }
                    }
                    steplog.End = DateTime.Now;
                    steplog.Outcome = true;
                }
                catch (Exception ex)
                {
                    steplog.End = DateTime.Now;
                    steplog.Ex = ex;
                    steplog.Outcome = false;
                    if (ex.InnerException != null)
                        System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
                    if (!ignoreerrors)
                        throw;
                }

            }
        }

        public void Populateworksheets(string filename, string defaultsheetname = "Unknown")
        {
            //open the file > create a stream
            FileStream stream = File.Open(filename, FileMode.Open, FileAccess.Read);
            Conditions condition = new Conditions();
            //create excel object
            IExcelDataReader excelReader;

            // Fix codepage issues
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            switch (condition.excelextension(filename))
            {
                case (int)Conditions.FileFormats.XLS:
                    //1.1 Reading from a binary Excel file ('97-2003 format; *.xls)
                    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                    break;
                case (int)Conditions.FileFormats.CSV:
                    //Handle CSV Files
                    excelReader = ExcelReaderFactory.CreateCsvReader(stream);
                    break;
                case (int)Conditions.FileFormats.XLSX:
                    //1.2 Reading from a OpenXml Excel file (2007 format; *.xlsx)
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    break;
                default:
                    throw new Exception("Unsupported extension: " + filename);
            }
            DataSet ds = excelReader.AsDataSet();
            //If table name is default (Normal for csv files)
            if (ds.Tables[0].TableName == "Table1")
            {
                ds.Tables[0].TableName = defaultsheetname;
            }

            //If it's the first time around populate first worksheet
            if (string.IsNullOrEmpty(firstworksheet))
            {
                //populate firstworksheet if it's null or blank
                firstworksheet = ds.Tables[0].TableName;

            }
            //Loop through the worksheets
            foreach (DataTable table in ds.Tables)
            {
                //First Worksheet with a given name take pre-emininence
                if (!worksheets.ContainsKey(table.TableName)) 
                    worksheets[table.TableName] = table;
            }

        }

        //register addon
        public void RegisterMethod(string action, Func<EasyExcelF, string[], bool> passedfunction)
        {
            //if null assign dict
            registeredactions ??= new Dictionary<string, Func<EasyExcelF, string[], bool>>(StringComparer.OrdinalIgnoreCase);
            registeredactions[action] = passedfunction;
        }
        public void calltestcase(string[] parms)
        {
            switch (parms[0].ToUpper())
            {
                case "CALL":
                    //call testcase in new scope
                    callnewtestcase(parms[1..]);
                    break;
                case "TEST":
                    //call testcase in new scope
                    test(parms[1..]);
                    break;
                default:
                    //call testcase in same scope
                    switch (parms.Length)
                    {
                        case 0:
                            throw new ArgumentOutOfRangeException("Expected Worksheet");
                        case 1:
                            this.Execute(parms[0]);
                            break;
                        default:
                            this.Execute(parms[0], parms[1..]);
                            break;
                    }
                    break;
            }

        }
        private void callnewtestcase(string[] parms)
        {
            EasyExcelF CalledTestcase = new EasyExcelF(this);
            switch (parms.Length)
            {
                case 0:
                    throw new ArgumentOutOfRangeException("Expected Worksheet");
                case 1:
                    CalledTestcase.Execute(parms[0]);
                    break;
                default:
                    CalledTestcase.Execute(parms[0], parms[1..]);
                    break;
            }

        }
        private void test(string[] parms)
        {
            if (StepHistory != null)
                throw new InvalidOperationException("Cannot Test within a test");
            EasyExcelF CalledTestcase = new EasyExcelF(this);
            CalledTestcase.StepHistory = new List<TestStepsLogEntry>();
            TestLog tl = new TestLog();
            tl.Started = DateTime.Now;
            tl.Test = parms[0];
            tl.parameters = parms;
            tl.Outcome = true;
            tl.StepHistory = CalledTestcase.StepHistory;
            try
            {
                switch (parms.Length)
                {
                    case 0:
                        throw new ArgumentOutOfRangeException("Expected Worksheet");
                    case 1:
                        CalledTestcase.Execute(parms[0]);
                        break;
                    default:
                        CalledTestcase.Execute(parms[0], parms[0..]);
                        break;
                }
            }
            catch (Exception ex)
            {
                tl.Outcome = false;
                tl.End = DateTime.Now;
                if (this.Screenshot != null)
                {
                    tl.StepHistory[tl.StepHistory.Count-1].screenshot = this.Screenshot(defaultpath);
                }
                tl.StepHistory[tl.StepHistory.Count - 1].Ex = ex;

                //this.TestHistory.Add(tl);
                //throw;
            }
            tl.End = DateTime.Now;
            this.TestHistory.Add(tl);

        }
        public void RegisterScreenShot(Func<string, string> sshot) => Screenshot = sshot;
        public void RegisterReport(Func<EasyExcelF, string> Report) => this.report = Report;
    }

}