
namespace EasyExcelFramework
{
    internal class EECore
    {
        public EECore(EasyExcelF ee)
        {
            //Initialise Addon
            ee.RegisterMethod("COMMENT", comments);
            ee.RegisterMethod("COMMENTS", comments);
            ee.RegisterMethod("COMMENT", comments);
            ee.RegisterMethod("STOP", stop);
            ee.RegisterMethod("LOAD FILE", loadfile);
            ee.RegisterMethod("SET LOCAL", setlocal);
            ee.RegisterMethod("SET LOCAL STRING", setlocal);
            ee.RegisterMethod("SET LOCAL DATETIME", setlocaldatetime);
            ee.RegisterMethod("SET LOCAL FLOAT", setlocalfloat);
            ee.RegisterMethod("SET LOCAL INT", setlocalint);
            ee.RegisterMethod("SET GLOBAL", setglobal);
            ee.RegisterMethod("SET GLOBAL STRING", setglobal);
            ee.RegisterMethod("SET GLOBAL DATETIME", setglobaldatetime);
            ee.RegisterMethod("SET GLOBAL FLOAT", setglobalfloat);
            ee.RegisterMethod("SET GLOBAL INT", setglobalint);
            ee.RegisterMethod("PARAMETERS", parameters);
            ee.RegisterMethod("PAUSE", pause);
            ee.RegisterMethod("SET RANDOM", setrandom);

        }
        private bool comments(EasyExcelF ee, string[] parms)
        {
            //Ignore Comments
            return true;
        }
        private bool pause(EasyExcelF ee, string[] parms)
        {
            //pause in seconds
            Thread.Sleep((int)(float.Parse(parms[0].ToString())*1000));
            return true;
        }
        private bool parameters(EasyExcelF ee, string[] parms)
        {
            for (int i = 0; i < parms.Length - 1; i++)
            {
                if (!string.IsNullOrEmpty(parms[i]))
                    ee.Locals[parms[i]] = ee.passedparams[i];
            }
            return true;
        }
        private bool stop(EasyExcelF ee, string[] parms)
        {
            //Stop test execution
            throw new Exception("Stop Called");
        }
        private bool loadfile(EasyExcelF ee, string[] parms)
        {
            //check for null
            if (ee.Worksheets[ee.Worksheet].Columns.Count - ee.CurrentIndent == 2)
            {
                //populate with just filenasme
                ee.Populateworksheets(parms[0].ToString());
            }
            else if (ee.Worksheets[ee.Worksheet].Columns.Count - ee.CurrentIndent > 2)
            {
                //set default sheet name
                ee.Populateworksheets(parms[0].ToString(), parms[1].ToString());
            }
            else
            {
                //Expected parameter
                throw new IndexOutOfRangeException("Expected Parameter");
            }

            return true;
        }
        private bool setlocal(EasyExcelF ee, string[] parms)
        {
            StringConverter sc = new StringConverter();
            //check it has a variable name and a value
            if (ee.Worksheets[ee.Worksheet].Columns.Count - ee.CurrentIndent < 2)
            {
                throw new IndexOutOfRangeException("Local Variable cannot be blank or null");
            }
            else
            {
                //assign variable
                try
                {
                    ee.Locals[parms[0].ToString()] = ee.Interpreter.EvalToString(ee, parms[1].ToString(), parms).ToString();
                }
                catch
                {
                    ee.Locals[parms[0].ToString()] = parms[1].ToString();
                }
            }

            return true;
        }
        private bool setlocaldatetime(EasyExcelF ee, string[] parms)
        {
            StringConverter sc = new StringConverter();
            //check it has a variable name and a value
            if (ee.Worksheets[ee.Worksheet].Columns.Count - ee.CurrentIndent < 2)
            {
                throw new IndexOutOfRangeException("Local Variable cannot be blank or null");
            }
            else
            {
                //assign variable
                try
                {
                    ee.Locals[parms[0].ToString()] = (DateTime)ee.Interpreter.DynamicEval(ee, parms[1].ToString(), parms);
                }
                catch
                {
                    ee.Locals[parms[0].ToString()] = parms[1].ToString();
                }
            }

            return true;
        }
        private bool setlocalint(EasyExcelF ee, string[] parms)
        {
            StringConverter sc = new StringConverter();
            //check it has a variable name and a value
            if (ee.Worksheets[ee.Worksheet].Columns.Count - ee.CurrentIndent < 2)
            {
                throw new IndexOutOfRangeException("Local Variable cannot be blank or null");
            }
            else
            {
                //assign variable
                try
                {
                    ee.Locals[parms[0].ToString()] = (int)ee.Interpreter.DynamicEval(ee, parms[1].ToString(), parms);
                }
                catch
                {
                    ee.Locals[parms[0].ToString()] = parms[1].ToString();
                }
            }

            return true;
        }
        private bool setlocalfloat(EasyExcelF ee, string[] parms)
        {
            StringConverter sc = new StringConverter();
            //check it has a variable name and a value
            if (ee.Worksheets[ee.Worksheet].Columns.Count - ee.CurrentIndent < 2)
            {
                throw new IndexOutOfRangeException("Local Variable cannot be blank or null");
            }
            else
            {
                //assign variable
                try
                {
                    ee.Locals[parms[0].ToString()] = (float)ee.Interpreter.DynamicEval(ee, parms[1].ToString(), parms);
                }
                catch
                {
                    ee.Locals[parms[0].ToString()] = parms[1].ToString();
                }
            }

            return true;
        }
        private bool setglobal(EasyExcelF ee, string[] parms)
        {
            StringConverter sc = new StringConverter();
            //check it has a variable name and a value
            if (ee.Worksheets[ee.Worksheet].Columns.Count - ee.CurrentIndent < 2)
                throw new IndexOutOfRangeException("Global Variable cannot be blank or null");
            //assign variable
            try
            {
                ee.Globals[parms[0].ToString()] = ee.Interpreter.EvalToString(ee, parms[1].ToString(), parms);

            }
            catch
            {
                ee.Globals[parms[0].ToString()] = parms[1].ToString();

            }
            return true;
        }
        private bool setglobaldatetime(EasyExcelF ee, string[] parms)
        {
            StringConverter sc = new StringConverter();
            //check it has a variable name and a value
            if (ee.Worksheets[ee.Worksheet].Columns.Count - ee.CurrentIndent < 2)
                throw new IndexOutOfRangeException("Global Variable cannot be blank or null");
            //assign variable
            try
            {
                ee.Globals[parms[0].ToString()] = (DateTime)ee.Interpreter.DynamicEval(ee, parms[1].ToString(), parms);

            }
            catch
            {
                ee.Globals[parms[0].ToString()] = parms[1].ToString();

            }
            return true;
        }
        private bool setglobalfloat(EasyExcelF ee, string[] parms)
        {
            StringConverter sc = new StringConverter();
            //check it has a variable name and a value
            if (ee.Worksheets[ee.Worksheet].Columns.Count - ee.CurrentIndent < 2)
                throw new IndexOutOfRangeException("Global Variable cannot be blank or null");
            //assign variable
            try
            {
                ee.Globals[parms[0].ToString()] = (float)ee.Interpreter.DynamicEval(ee, parms[1].ToString(), parms);

            }
            catch
            {
                ee.Globals[parms[0].ToString()] = parms[1].ToString();

            }
            return true;
        }
        private bool setglobalint(EasyExcelF ee, string[] parms)
        {
            StringConverter sc = new StringConverter();
            //check it has a variable name and a value
            if (ee.Worksheets[ee.Worksheet].Columns.Count - ee.CurrentIndent < 2)
                throw new IndexOutOfRangeException("Global Variable cannot be blank or null");
            //assign variable
            try
            {
                ee.Globals[parms[0].ToString()] = (int)ee.Interpreter.DynamicEval(ee, parms[1].ToString(), parms);

            }
            catch
            {
                ee.Globals[parms[0].ToString()] = parms[1].ToString();

            }
            return true;
        }

        private bool setrandom(EasyExcelF ee, string[] parms)
        {
            //check it has a variable name and a value
            if (ee.Worksheets[ee.Worksheet].Columns.Count - ee.CurrentIndent < 2)
            {
                throw new IndexOutOfRangeException("Local Variable cannot be blank or null");
            }
            else
            {
                Random rand = new Random();
                //assign variable
                try
                {
                    int minval = int.Parse(parms[1]);
                    int maxval = int.Parse(parms[2]);
                    ee.Locals[parms[0].ToString()] = (int)rand.Next(minval, maxval);
                }
                catch
                {
                    int maxval = int.Parse(parms[1]);
                    ee.Locals[parms[0].ToString()] = (int)rand.Next(maxval);
                }
            }

            return true;
        }
    }
}
