using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace laba_xl
{
    class Parser
    {
        private const string ERR_PARENS = "Пропущено закриваючу дужку";
        private const string ERR_OPER_EXPRESSION = "Оператор чи вираз після нього відсутні або в некоректному форматі";
        private const string ERR_ZERO_DIVISION = "Неприпустима дія: ділення на нуль";
        private const string ERR_OPER_ABSENT = "Знак оператора після числа або дужок відсутній або неприпустимий";
        private const string ERR_RECURSION = "Виявлено рекурсивне посилання комірки на саму себе";
        private const string ERR_CELL_DOESNT_EXIST = "Комірка, посилання на яку ви ввели, не існує";
        private const string ERR_EXTRA_PARENS = "Виявлено зайву закриваючу дужку";
        private const string ERR_BEFORE_POINT = "перед крапкою має бути число";
        private Token _currentToken;
        private Tokenizer _tokenizer;

        private string _error;
        private string _invokerName;

        private int _lparenCount = 0;
        private int _rparenCount = 0;

        public string Error { get { return _error; } }
        public string InvokerName { get { return _invokerName; } }

        // Initializes a new Parser with the provided tokenizer
        // and the invoker's - cell's whose expression parsing is required - name.
        public Parser(Tokenizer tokenizer, string invokerName)
        {
            _tokenizer = tokenizer;
            _currentToken = _tokenizer.ReadNextToken();
            _error = "";
            _invokerName = invokerName;
        }

        // Parses next token in the given expression that might, however, be erronous.
        private bool ParseNextToken(TokenType tokenType)
        {
            if (_currentToken.Type == tokenType)
            {
                return CheckTokenCorrectness(tokenType);
            }
            else
            {
                _error = tokenType == TokenType.RPAREN ? ERR_PARENS : ERR_OPER_EXPRESSION;
                return false;
            }
        }

        // Checks whether or not the current token is NOT erronous.
        private bool CheckTokenCorrectness(TokenType tokenType)
        {
            if (!CheckParenthesesCorrectness(tokenType))
            {
                return false;
            }

            _currentToken = _tokenizer.ReadNextToken();

            if ((tokenType == TokenType.RPAREN || tokenType == TokenType.NUMBER)
            && (_currentToken.Type == TokenType.LPAREN || _currentToken.Type == TokenType.NUMBER))
            {
                _error = ERR_OPER_ABSENT;
                return false;
            }

            return (_error = _tokenizer.Error).Equals("");
        }

        // Checks the correspondence of opening/closing parentheses' amounts.
        private bool CheckParenthesesCorrectness(TokenType tokenType)
        {
            if (tokenType == TokenType.LPAREN)
            {
                _lparenCount++;
            }
            if (tokenType == TokenType.RPAREN)
            {
                _rparenCount++;
                if (_rparenCount == _lparenCount)
                {
                    if (_tokenizer.CurrentChar == ')')
                    {
                        _error = ERR_EXTRA_PARENS;
                        return false;
                    }
                    _lparenCount = 0;
                    _rparenCount = 0;
                }
            }

            return true;
        }

        // Parses an expression with possibly present comparison(s).
        public Node ParseExpression()
        {
            Node leftSide = ParseComparand();

            if (leftSide == null)
            {
                return null;
            }

            return LoopComparison(leftSide);
        }

        // Loops through the comparison's right side (if present).
        private Node LoopComparison(Node node)
        {
            Node result = node;
            Node leftSide = result;
            return result;
        }


        // Parses the comparand with possibly present addition/subtraction.
        private Node ParseComparand()
        {
            Node leftSide = ParseAddend();

            if (leftSide == null)
            {
                return null;
            }

            return LoopAdditionSubtraction(leftSide);
        }

        // Loops through the addition/subtraction's right side (if present).
        private Node LoopAdditionSubtraction(Node leftSide)
        {
            while (_currentToken.Type == TokenType.PLUS || _currentToken.Type == TokenType.MINUS)
            {
                Token operation = _currentToken;
                ParseNextToken(operation.Type);

                Node rightSide = ParseAddend();
                if (rightSide == null)
                {
                    return null;
                }

                leftSide = new ArithmeticOperationNode(leftSide, operation, rightSide);
            }
            return leftSide;
        }

        // Parses the addend with possibly present multiplication/(integer) division/modulo.
        private Node ParseAddend()
        {
            Node leftSide = ParseNegatand();

            if (leftSide == null)
            {
                return null;
            }

            return LoopFactorization(leftSide);
        }

        // Loops through the multiplication/(integer) division/modulo's right side (if present).
        private Node LoopFactorization(Node leftSide)
        {
            while (_currentToken.Type == TokenType.MULTIPLY || _currentToken.Type == TokenType.DIV
            || _currentToken.Type == TokenType.MOD || _currentToken.Type == TokenType.DIVIDE || _currentToken.Type == TokenType.POW)
            {
                Token operation = _currentToken;
                ParseNextToken(operation.Type);

                Node rightSide = ParseNegatand();
                if (rightSide == null)
                {
                    return null;
                }

                if ((operation.Type == TokenType.DIVIDE || operation.Type == TokenType.DIVIDE) && rightSide.GetValue() == 0)
                {
                    _error = ERR_ZERO_DIVISION;
                    return null;
                }

                leftSide = new ArithmeticOperationNode(leftSide, operation, rightSide);
            }
            return leftSide;
        }

        // Parses the negatand with possibly present inner negations.
        private Node ParseNegatand()
        {
            return ParseIncDec();
        }

        // Parses the inc/dec function calls if present, otherwise parses primary expression.
        // NB: the inc/dec with cell reference as the argument
        // will only modify the corresponding cell's value if:
        // a) its formula is a certain number, but NOT an expression
        // b) you invoke it from the just edited cell, unlike when all cell values are updated.
        private Node ParseIncDec()
        {
            Node result;
            Token operation = _currentToken;
            if (operation.Type == TokenType.INC || operation.Type == TokenType.DEC)
            {
                if (!ParseNextToken(operation.Type))
                {
                    return null;
                }

                // Parse as regular parentheses expression
                result = ParseParenthesesExpression();

                return result == null ? null : new IncDecNode(result, operation.Type, _invokerName);
            }

            return ParsePrimary();
        }

        // Parses the primary expression,
        // which can be either a number, a cell reference or the expression in parentheses.
        private Node ParsePrimary()
        {
            // Syntax error occured in the tokenizer.
            if (!_tokenizer.Error.Equals(""))
            {
                _error = _tokenizer.Error;
                return null;
            }

            return ParseAtom();
        }

        // Parses the primary (atomic) expression depending on what it is.
        private Node ParseAtom()
        {
            Token token = _currentToken;

            if (token.Type == TokenType.NUMBER)
            {
                return ParseNextToken(TokenType.NUMBER) ? new NumberNode(token) : null;
            }

            if (token.Type == TokenType.CELL)
            {
                return ParseCell(token.Value);
            }

            if (token.Type == TokenType.LPAREN)
            {
                Node result = ParseParenthesesExpression();

                return result;
            }

            _error = ERR_OPER_EXPRESSION;
            return null;
        }

        // Parses the cell reference.
        private Node ParseCell(string reference)
        {
            Cell cell = CellManager.Instance.GetCell(reference);
            if (cell == null)
            {
                _error = ERR_CELL_DOESNT_EXIST;
                return null;
            }

            if (CellManager.Instance.HasReferenceRecursion(cell, _invokerName))
            {
                _error = ERR_RECURSION;
                return null;
            }

            AddCellToReferencesList(cell);
            return ParseNextToken(TokenType.CELL) ? new CellNode(reference) : null;
        }

        // Adds the referenced cell to the current one's reference list
        // for further reference recursion checks.
        private void AddCellToReferencesList(Cell cell)
        {
            Cell invoker = CellManager.Instance.GetCell(_invokerName);
            Cell currentCell = CellManager.Instance.CurrentCell;
            if (invoker.Name.Equals(currentCell.Name) && !invoker.CellReferences.Contains(cell))
            {
                invoker.CellReferences.Add(cell);
            }
        }

        // Parses the expression in parentheses (here we go again...)
        private Node ParseParenthesesExpression()
        {
            ParseNextToken(TokenType.LPAREN);
            Node node = ParseExpression();

            if (node == null)
            {
                return null;
            }


            return ParseNextToken(TokenType.RPAREN) ? node : null;
        }
    }
}

