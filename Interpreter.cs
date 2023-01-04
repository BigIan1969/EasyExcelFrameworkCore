using System.Collections;

namespace EasyExcelFramework
{
    public class InterpreterClass
    {
        public bool Eval(EasyExcelF ee, string expression, dynamic[] parms)
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
            int paramno = 0;
            foreach (string param in parms)
            {
                paramno++;
                DExpresso.SetVariable("PARAM" + paramno, parms[paramno - 1]);
            }
            return (bool)DExpresso.Eval(expression);

        }
        public bool Eval(EasyExcelF ee, string expression)
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
        public string EvalToString(EasyExcelF ee, string expression, dynamic[] parms)
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
            int paramno = 0;
            foreach (string param in ee.passedparams)
            {
                paramno++;
                DExpresso.SetVariable("PARAM" + paramno, param);
            }
            return (string)DExpresso.Eval(expression).ToString();

        }
        public dynamic DynamicEval(EasyExcelF ee, string expression, dynamic[] parms)
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
            StringConverter sc=new StringConverter();
            int paramno = 0;
            foreach (string param in ee.passedparams)
            {
                paramno++;
                DExpresso.SetVariable("PARAM" + paramno, sc.DetectType(param));
            }
            return DExpresso.Eval(expression);

        }
    }
}
