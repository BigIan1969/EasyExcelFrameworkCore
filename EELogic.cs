using System.Data;
using System.Reflection;

namespace EasyExcelFramework
{
    internal class EELogic
    {
        //Logic helpers
        public bool ElseActive;
        public string SwitchVal;
        public EELogic(EasyExcelF ee)
        {
            ee.RegisterMethod("IF", localif);
            ee.RegisterMethod("ELSE", localelse);
            ee.RegisterMethod("SWITCH", localswitch);
            ee.RegisterMethod("CASE", localcase);
            ee.RegisterMethod("LOOP", loop);
            ee.RegisterMethod("BREAK", localbreak);
        }
        private bool localif(EasyExcelF ee, string[] parms)
        {
            if (ee.Interpreter.Eval(ee, parms[0].ToString(), parms[1..]))
            {
                ElseActive = false;
                ee.calltestcase(parms[1..]);
            }
            else
            {
                ElseActive = true;
            }
            return true;
        }
        private bool localelse(EasyExcelF ee, string[] parms)
        {
            if (ElseActive)
            {
                ElseActive = false;
                SwitchVal = "";
                ee.calltestcase(parms);
            }
            return true;
        }
        private bool localswitch(EasyExcelF ee, string[] parms)
        {
            try
            {
                SwitchVal = ee.Interpreter.EvalToString(ee, parms[0], parms[1..]).ToString();

            }
            catch
            {
                SwitchVal = parms[0];
            }
            return true;
        }
        private bool localcase(EasyExcelF ee, string[] parms)
        {
            if (SwitchVal == parms[0])
            {
                ee.calltestcase(parms[1..]);
            }
            else
            {
                ElseActive = true;
            }
            return true;
        }
        private bool loop(EasyExcelF ee, string[] parms)
        {
            if (ee.Worksheets.ContainsKey(parms[0]) & ee.Worksheets.ContainsKey(parms[1]))
            {
                ee.loopactive = true;
                foreach(DataRow r in ee.Worksheets[parms[0]].Rows)
                {
                    // Execute Worksheet for each row
                    // convert ItemArray
                    string[] rowdata = Array.ConvertAll(r.ItemArray, x => x.ToString());
                    // instanciate strings topass
                    string[] topass = new string[rowdata.Length+1];
                    //copy itemarray data to one index over
                    rowdata.CopyTo(topass, 1);
                    // set first item to worksheet to call
                    topass[0] = parms[1];

                    //call the worksheet
                    ee.calltestcase(topass);
                    if (ee.breakoutofloop)
                    {
                        ee.breakoutofloop = false;
                        ee.loopactive = false;
                        break;
                    }
                }
            }
            else if (ee.Worksheets.ContainsKey(parms[1]))
            {
                ee.loopactive = true;
                int iter;
                try
                {
                    iter = int.Parse(parms[0]);
                }
                catch
                {
                    throw new Exception("Cannot parse loop");
                }
                for (int i = 0; i < iter; i++)
                {
                    // Execute Worksheet for each iteration
                    ee.calltestcase(new string[] { "CALL", parms[1]});
                    if (ee.breakoutofloop)
                    {
                        ee.breakoutofloop = false;
                        ee.loopactive = false;
                        break;
                    }
                }
            }
            return true;
        }
        private bool localbreak(EasyExcelF ee, string[] parms)
        {
            if (ee.loopactive)
            {
                ee.breakoutofloop = true;
            }
            else
            {
                if (ee.parent!=null)
                {
                    _ = localbreak(ee.parent, parms);
                }
            }
            return true;
        }
    }
}
