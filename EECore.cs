using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyExcelFramework
{
    internal class EECore
    {
        public EECore(EasyExcelF ee)
        {
            //Initialise Addon
            ee.RegisterMethod("COMMENT", comments);
            ee.RegisterMethod("COMMENTS", comments);
            ee.RegisterMethod("STOP", stop);
            ee.RegisterMethod("LOAD FILE", loadfile);
            ee.RegisterMethod("SET LOCAL", setlocal);
            ee.RegisterMethod("SET GLOBAL", setglobal);
        }
        private bool comments(EasyExcelF ee, string[] parms)
        {
            //Ignore Comments
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
                    ee.Locals[parms[0].ToString()] = (string)InterpreterClass.EvalToString(ee, parms[1].ToString(),parms).ToString();
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
            //check it has a variable name and a value
            if (ee.Worksheets[ee.Worksheet].Columns.Count - ee.CurrentIndent < 2)
                throw new IndexOutOfRangeException("Global Variable cannot be blank or null");
            //assign variable
            try
            {
                ee.Globals[parms[0].ToString()] = (string)InterpreterClass.EvalToString(ee, parms[1].ToString(), parms);

            }
            catch
            {
                ee.Globals[parms[0].ToString()] = parms[1].ToString();

            }
            return true;
        }
    }
}
