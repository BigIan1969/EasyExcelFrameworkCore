using System.Collections;

namespace EasyExcelFramework
{
    public class InterpreterClass
    {
        public bool Eval(EasyExcelF ee, string expression, string[] parms)
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
        public string EvalToString(EasyExcelF ee, string expression, string[] parms)
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
