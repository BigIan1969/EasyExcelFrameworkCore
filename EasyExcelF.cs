﻿using ExcelDataReader;
using System.Collections;
using System.Data;
using System.Runtime.CompilerServices;

namespace EasyExcelFramework
{
    public class EasyExcelF
    {
        //Dictionary of Worksheets
        private Dictionary<string, DataTable>? worksheets;
        public Dictionary<string, DataTable>? Worksheets { get => worksheets;}

        //worksheet property
        private string worksheet;
        public string Worksheet { get => worksheet; }

        //Dictionary of variables
        public Dictionary<string, string>? Locals;

        //Dictionary of Globals
        public Dictionary<string, string>? Globals;

        public IDictionary? Environ;
        //First worksheet where the framework begins executing

        //RegisteredActions
        private Dictionary<string, Func<EasyExcelF, string[], bool>> registeredactions;

        private string firstworksheet;

        //Current indent the framework is operating at
        private int currentindent;
        public int CurrentIndent { get => currentindent; }

        private string[] currentdatarow;

        public string[] Currentdatarow { get => currentdatarow;  }
        public int Currentnownumber { get => currentnownumber; }

        private int currentnownumber;

        //Logic helpers
        public bool ElseActive;
        public string SwitchVal;

        public EasyExcelF(string filename = "default.xlsx")
        {
            //Add core
            EECore eec = new EECore(this);
            EELogic eel = new EELogic(this);

            //Instanciate worksheets
            worksheets = new Dictionary<string, DataTable>(StringComparer.OrdinalIgnoreCase);

            //get environment variables
            Environ= Environment.GetEnvironmentVariables();
            
            //set indent level
            currentindent = 0;
            //Instanciate variables
            Locals = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            Populateworksheets(filename);
            if (string.IsNullOrEmpty(firstworksheet))
                throw (new Exception("First worksheet not found: " + filename));
            //Execute(firstworksheet, new string[1]);
        }
        public EasyExcelF()
        {

            //Add core
            EECore eec = new EECore(this);
            EELogic eel = new EELogic(this);

            //Instanciate worksheets
            worksheets = new Dictionary<string, DataTable>(StringComparer.OrdinalIgnoreCase);
            currentindent = 0;

            //Instanciate variables
            Locals = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            Populateworksheets("Default.xlsx");
            if (string.IsNullOrEmpty(firstworksheet))
                throw (new Exception("First worksheet not found: Default.xslx"));
            //Execute(firstworksheet, new string[1]);
        }
        public EasyExcelF(EasyExcelF parent)
        {
            registeredactions = parent.registeredactions;

            //Instanciate worksheets
            worksheets = parent.worksheets;
            currentindent = 0;
            Globals = parent.Globals;

            //Instanciate variables
            Locals = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        }
        public void Execute(string worksheet = null, string[] passedparameters = null)
        {

            //handle null parameter
            if (worksheet is null)
            {
                worksheet = firstworksheet;
            }
            if (passedparameters is null)
            {
                passedparameters = new string[1];
            }

            //handle null worksheets
            if (worksheets == null)
            {
                throw new ArgumentNullException(nameof(worksheets));
            }
            if (Globals == null)
                Globals = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            this.worksheet = worksheet;

            this.currentnownumber = 0;

            //loop through specified worksheet
            foreach (DataRow execrow in worksheets[worksheet].Rows)
            {
                this.currentnownumber++;
                if (execrow[currentindent] == null)
                    throw new ArgumentNullException(nameof(execrow));
                
                //Convert ItemArray into string array
                this.currentdatarow = Array.ConvertAll(execrow.ItemArray, x => x.ToString());

                int ind = this.currentindent + 1;
                string[] parms = this.currentdatarow[ind..];

                if (string.IsNullOrEmpty(SwitchVal))
                { //Process Normally
                    if (registeredactions.ContainsKey(execrow[currentindent].ToString()))
                    {
                        bool result = registeredactions[execrow[currentindent].ToString()](this, parms);
                    }
                    else
                    {
                        //If it's a worksheet
                        if (worksheets.ContainsKey(execrow[0 + currentindent].ToString()))
                        {
                            Execute(execrow[0].ToString(), parms);

                        }
                        else
                        {
                            //assign variable
                            Locals[execrow[0 + currentindent].ToString()] = execrow[1 + currentindent].ToString();
                        }
                    }
                }
                else
                { //process switch
                    if (execrow[currentindent].ToString() == SwitchVal)
                    {
                        ind += 2;
                        EasyExcelF CalledTestcase = new EasyExcelF(this);
                        CalledTestcase.Execute(execrow[currentindent+1].ToString());
                        CalledTestcase = null;
                    }
                    else if (execrow[currentindent].ToString().ToUpper()=="ELSE")
                    {
                        ind += 2;
                        ElseActive = false;
                        EasyExcelF CalledTestcase = new EasyExcelF(this);
                        CalledTestcase.Execute(execrow[currentindent + 1].ToString());
                        CalledTestcase = null;
                    }
                    else
                    {
                        ElseActive = true;
                    }
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
                worksheets[table.TableName] = table;
            }

        }

        //register addon
        public void RegisterMethod(string action , Func<EasyExcelF , string[], bool> passedfunction)
        {
            //if null assign dict
            registeredactions ??= new Dictionary<string, Func<EasyExcelF, string[], bool>>(StringComparer.OrdinalIgnoreCase);
            registeredactions[action]=passedfunction;
        }

    }

}