using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        }
        private bool localif(EasyExcelF ee, string[] parms)
        {
            if (InterpreterClass.Eval(ee, parms[0].ToString(), parms[1..]))
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
                SwitchVal = InterpreterClass.EvalToString(ee, parms[0], parms[1..]).ToString();

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

    }
}
