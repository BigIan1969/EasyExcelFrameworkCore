using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using DynamicExpresso;

namespace EasyExcelFramework
{
    internal static class InterpreterClass
    {
        public static bool Eval(EasyExcelF ee, string expression, string[] parms)
        {
            DynamicExpresso.Interpreter DExpresso = new DynamicExpresso.Interpreter();
            foreach (DictionaryEntry variable in ee.Environ)
            {
                DExpresso.SetVariable(variable.Key.ToString(), variable.Value);
            }
            foreach (var variable in ee.Globals)
            {
                DExpresso.SetVariable(variable.Key.ToString(), variable.Value);
            }
            foreach (var variable in ee.Locals)
            {
                DExpresso.SetVariable(variable.Key.ToString(), variable.Value);
            }
            return (bool)DExpresso.Eval(expression);

        }
        public static string EvalToString(EasyExcelF ee, string expression, string[] parms)
        {
            DynamicExpresso.Interpreter DExpresso = new DynamicExpresso.Interpreter();
            foreach (DictionaryEntry variable in ee.Environ)
            {
                DExpresso.SetVariable(variable.Key.ToString(), variable.Value);
            }
            foreach (var variable in ee.Globals)
            {
                DExpresso.SetVariable(variable.Key.ToString(), variable.Value);
            }
            foreach (var variable in ee.Locals)
            {
                DExpresso.SetVariable(variable.Key.ToString(), variable.Value);
            }
            return (string)DExpresso.Eval(expression).ToString();

        }
    }
}
