using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace laba_xl
{
    public enum TokenType
    {
        EOX, // End of expression
        NUMBER, // 0123456789
        PLUS, // <...> + <...>
        MINUS, // <...> - <...>
        MULTIPLY, // <...> * <...>
        DIVIDE, // <...> / <...> (common division)
        MOD, // <...> % <...>
        DIV, // <...> : <...> (integer division)
        POW, // <...> ^ <...>
        LPAREN, // (
        RPAREN, // )
        DEC, // DEC(<...>): <...> = <...> - 1
        INC, // INC(<...>): <...> = <...> + 1
        CELL, // RX:CY cell reference, where X and Y - some indices in the data grid view
        ERROR // Erronous token
    }

    public class Token
    {
        private TokenType _type;
        private string _value;

        public TokenType Type { get { return _type; } }
        public string Value { get { return _value; } }

        public Token(TokenType type, string val)
        {
            _type = type;
            _value = val;
        }
    }

    public class Tokenizer
    {
        private const char NIX = '\0';
        private const char PLUS = '+';
        private const char MINUS = '-';
        private const char MULTIPLY = '*';
        private const char DIVIDE = '/';
        private const char MOD = '%';
        private const char DIV = ':';
        private const char POW = '^';
        private const char LPAREN = '(';
        private const char RPAREN = ')';

        private const string DEC = "DEC";
        private const string INC = "INC";

        private const string ERR_MANY_POINTS = "У введеному числі більше однієї десяткової крапки";
        private const string ERR_LETTERS = "Введено літери після цифр";
        private const string ERR_CELL = "Некоректний формат імені комірки";
        private const string ERR_INVALID_TOKEN = "Введено некоректний символ(и)";
        private const string ERR_AFTER_POINT = "Відсутні цифри після крапки";
        private const string ERR_BEFORE_POINT = "перед крапкою має бути число";
        private const string oksymbols = "()1234567890*/^:+-%";

        private List<char> _singleChars = new List<char>()
{ PLUS, MINUS, MULTIPLY, DIVIDE, MOD, DIV, LPAREN, RPAREN, POW };

        private Dictionary<char, TokenType> _tokenPairs = new Dictionary<char, TokenType>()
        {
            [PLUS] = TokenType.PLUS,
            [MINUS] = TokenType.MINUS,
            [MULTIPLY] = TokenType.MULTIPLY,
            [DIVIDE] = TokenType.DIVIDE,
            [MOD] = TokenType.MOD,
            [DIV] = TokenType.DIV,
            [LPAREN] = TokenType.LPAREN,
            [RPAREN] = TokenType.RPAREN,
            [POW] = TokenType.POW,
        };

        private string _expression;
        private char _currentChar;
        private int _pos;
        private string _error;

        public string Error { get { return _error; } }
        public char CurrentChar { get { return _currentChar; } }

        public Tokenizer(string expression)
        {
            _expression = expression;
            _error = "";
            _pos = 0;
            _currentChar = expression[0];
        }

        // Reads next character in the expression.
        private void ReadNextChar()
        {
            _pos++;
            _currentChar = _pos > _expression.Length - 1 ? NIX : _expression[_pos];
        }

        // If a user enters whitespaces, they need to be skipped using this method.
        private void SkipWhiteSpace()
        {
            while (_currentChar != NIX && Char.IsWhiteSpace(_currentChar))
            {
                ReadNextChar();
            }
        }

        // Reads next token in the expression.
        public Token ReadNextToken()
        {
            while (_currentChar != NIX)
            {
                SkipWhiteSpace();

                if (Char.IsLetter(_currentChar) || !oksymbols.Contains(_currentChar))
                {
                    return ResolveLetters();
                }

                if (Char.IsDigit(_currentChar))
                {
                    return ResolveNumber();
                }

                if (_singleChars.Contains(_currentChar))
                {
                    return ResolveSingleCharOperation();
                }
            }
            return new Token(TokenType.EOX, NIX.ToString());
        }

        // Resolve a token starting with letter
        // (may be either a cell, a dec/inc function or an erronous token).
        private Token ResolveLetters()
        {
            char firstChar = Char.ToUpper(_currentChar);

            if (firstChar == 'D' || firstChar == 'I')
            {
                return ResolveIncDec(firstChar);
            }

            return ResolveCell();
        }

        // Resolves an inc/dec function or an erronous token
        // if there are erronous characters after alleged DEC/INC first characters.
        private Token ResolveIncDec(char firstChar)
        {
            if (!ReadIncDec(firstChar, out string charSequence))
            {
                return ResolveErronousToken(ERR_INVALID_TOKEN);
            }

            if (charSequence.Equals(INC))
            {
                return new Token(TokenType.INC, INC);
            }

            if (charSequence.Equals(DEC))
            {
                return new Token(TokenType.DEC, DEC);
            }

            return ResolveErronousToken(ERR_INVALID_TOKEN);
        }

        // Tries to read the "inc"/"dec" notation in the expression.
        // If fails, an erronous token will be returned hereinafter.
        private bool ReadIncDec(char firstChar, out string charSequence)
        {
            charSequence = firstChar.ToString();
            char current;
            int i = 1;

            while ((INC.Contains(charSequence) || DEC.Contains(charSequence)) && i < INC.Length)
            {
                ReadNextChar();
                current = Char.ToUpper(_currentChar);

                if (!(charSequence.Equals(INC.Substring(0, i)) || charSequence.Equals(DEC.Substring(0, i))))
                {
                    return false;
                }

                charSequence += current.ToString();
                i++;
            }
            ReadNextChar();

            return true;
        }

        // Resolves a cell reference or an erronous token
        // if the token is not of RXCY format.
        private Token ResolveCell()
        {
            string result = "";

            while (_currentChar != NIX && Char.IsLetterOrDigit(_currentChar))
            {
                result += _currentChar;
                ReadNextChar();
            }

            var matches = new Regex(@"^R(?<row>\d+)C(?<col>\d+)$").Matches(result);

            if (matches.Count != 1)
            {
                return ResolveErronousToken(ERR_CELL);
            }
            return new Token(TokenType.CELL, matches[0].Groups[0].Value);
        }

        // Resolves a real number.
        private Token ResolveNumber()
        {
            Token result = ResolvePoint();
            return result;
        }

        // Resolves a number's inner structure.
        // If it contains multiple decimal points or letters, an erronous token is returned.
        private Token ResolvePoint()
        {
            string value = "";

            while (_currentChar != NIX && (Char.IsDigit(_currentChar)))
            {
                value += _currentChar;
                ReadNextChar();
                if (Char.IsLetter(_currentChar))
                {
                    return ResolveErronousToken(ERR_LETTERS);
                }
            }
            return new Token(TokenType.NUMBER, value);
        }


        // Resolves single-charactered operations, such as "+", "-", etc.
        private Token ResolveSingleCharOperation()
        {
            char tmp = _currentChar;
            ReadNextChar();
            return new Token(_tokenPairs[tmp], tmp.ToString());
        }

        // Resolves an erronous token which also contains error info.
        private Token ResolveErronousToken(string error)
        {
            _error = error;
            return new Token(TokenType.ERROR, error);
        }
    }
}



