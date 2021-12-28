using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace laba_xl
{
    abstract class Node
    {
        public abstract double GetValue();
    }

    class NumberNode : Node
    {
        private double _value;
        private Token _token;

        public NumberNode(Token token)
        {
            _token = token;
            _value = double.Parse(token.Value);
        }

        public override double GetValue()
        {
            return _value;
        }
    }

    class ArithmeticOperationNode : Node
    {
        private Node _leftSide;
        private Node _rightSide;
        private Token _operation;

        public ArithmeticOperationNode(Node leftSide, Token operation, Node rightSide)
        {
            _leftSide = leftSide;
            _operation = operation;
            _rightSide = rightSide;
        }

        public override double GetValue()
        {
            switch (_operation.Type)
            {
                case TokenType.PLUS:
                    return _leftSide.GetValue() + _rightSide.GetValue();
                case TokenType.MINUS:
                    return _leftSide.GetValue() - _rightSide.GetValue();
                case TokenType.MULTIPLY:
                    return _leftSide.GetValue() * _rightSide.GetValue();
                case TokenType.DIVIDE:
                    return _leftSide.GetValue() / _rightSide.GetValue();
                case TokenType.MOD:
                    return _leftSide.GetValue() % _rightSide.GetValue();
                case TokenType.POW:
                    return Math.Pow(_leftSide.GetValue(), _rightSide.GetValue());
                case TokenType.DIV:
                    return (int)(_leftSide.GetValue() / _rightSide.GetValue());
                default:
                    return 0;
            }
        }
    }

    class IncDecNode : Node
    {
        private Node _node;
        private TokenType _operation;
        private string _invokerName;

        public IncDecNode(Node node, TokenType operation, string invokerName)
        {
            _node = node;
            _operation = operation;
            _invokerName = invokerName;
        }

        public override double GetValue()
        {
            if (_node is CellNode)
            {
                CellNode cellNode = (CellNode)_node;

                if (_invokerName.Equals(CellManager.Instance.CurrentCell.Name))
                {
                    cellNode.AddedValue += _operation == TokenType.INC ? 1 : -1;
                    return cellNode.GetValue();
                }
            }

            return _operation == TokenType.INC ? _node.GetValue() + 1 : _node.GetValue() - 1;
        }
    }

    class CellNode : Node
    {
        private string _name;
        private Cell _cell;

        public int AddedValue { get; set; }

        public CellNode(string name)
        {
            AddedValue = 0;
            _name = name;
            _cell = CellManager.Instance.GetCell(_name);
        }

        public override double GetValue()
        {
            double result = Convert.ToDouble(_cell.Value) + AddedValue;
            _cell.Value = result.ToString();
            return result;
        }
    }
}
