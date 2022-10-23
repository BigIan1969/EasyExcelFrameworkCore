using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace EasyExcelFramework
{
    internal class EEAssert
    {
        public EEAssert(EasyExcelF ee)
        {
            //Initialise Addon
            ee.RegisterMethod("ASSERT IF", assertif);

        }
        private bool assertif(EasyExcelF ee, string[] parms)
        {
            if (parms.Length==1)
                Assert.IsTrue(ee.Interpreter.Eval(ee, parms[0], parms));
            else
                Assert.IsTrue(ee.Interpreter.Eval(ee, parms[0], parms), parms[1]);
            return true;
        }
    }
}
