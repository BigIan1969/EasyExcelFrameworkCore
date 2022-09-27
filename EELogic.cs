using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyExcelFramework
{
    internal class EELogic
    {
        public EELogic(EasyExcelF ee)
        {
            ee.RegisterMethod("IF", localif);
            ee.RegisterMethod("ELSE", localelse);
            ee.RegisterMethod("SWITCH", localswitch);
        }
        private bool localif(EasyExcelF ee, string[] parms)
        {
            if (InterpreterClass.Eval(ee, parms[0].ToString(), parms[1..]))
            {
                ee.ElseActive = false;
                EasyExcelF CalledTestcase = new EasyExcelF(ee);
                CalledTestcase.Execute(parms[1], parms[2..]);
            }
            else
            {
                ee.ElseActive = true;
            }
            return true;
        }
        private bool localelse(EasyExcelF ee, string[] parms)
        {
            if (ee.ElseActive)
            {
                ee.ElseActive = false;
                ee.SwitchVal = "";
                EasyExcelF CalledTestcase = new EasyExcelF(ee);
                CalledTestcase.Execute(parms[0], parms[1..]);
            }
            else
            {
                throw new InvalidOperationException("Enexpected Else");
            }
            return true;
        }
        private bool localswitch(EasyExcelF ee, string[] parms)
        {
            try
            {
                ee.SwitchVal = InterpreterClass.EvalToString(ee, parms[0], parms[1..]).ToString();

            }
            catch
            {
                ee.SwitchVal = parms[0];
            }
            return true;
        }
    }
}
