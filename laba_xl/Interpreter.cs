using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace laba_xl
{
    class Interpreter
    {
        private Parser _parser;
        private string _error;

        public string Error { get { return _error; } }

        public Interpreter(string expression, string invokerName)
        {
            _error = "";
            var _tokenizer = expression.Equals("") ? new Tokenizer("0") : new Tokenizer(expression);
            _parser = new Parser(_tokenizer, invokerName);
        }

        public string EvaluateExpression()
        {
            Node resultingTree = _parser.ParseExpression();

            if (resultingTree == null)
            {
                _error = _parser.Error;
                return _error;
            }

            return resultingTree.GetValue().ToString();
        }
    }
}
