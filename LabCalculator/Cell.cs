using Antlr4.Runtime;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace LabCalculator
{
    class Cell
    {

        private DataGridViewCell _parent;
        private string _value;
        private string _formula;


        private string _error;
        private string _name;


        public string Value
        {
            get
            {
                return _value;
            }
            set
            {
                _value = value.Equals("0") ? "0" : "1";
            }
        }

            public string Formula
        {
            get { return _formula; }
            set { _formula = value; }
        }

        public DataGridViewCell Parent {  get { return _parent; } }

        public string Name { get { return _name; } }

        public string Error
        {
            get { return _error; }
            set { _error = value; }
        }

        public List<Cell> CellReferences { get; set; }

        public Cell(DataGridViewCell parent, string name, string formula)
        {
            _parent = parent;
            _name = name;
            _formula = "0";
            _error = "";
            CellReferences = new List<Cell>();
            CellReferences.Add(new Cell());
        }

        public Cell()
        {
            _name = "";
        }


        public double Evaluate()
        {
            var lexer = new LabCalculatorLexer(new AntlrInputStream(Formula));
            lexer.RemoveErrorListeners();
            lexer.AddErrorListener(new ThrowExceptionErrorListener());

            var tokens = new CommonTokenStream(lexer);
            var parser = new LabCalculatorParser(tokens);

            var tree = parser.compileUnit();

            var visitor = new LabCalculatorVisitor();

            return visitor.Visit(tree);
        }


    }


    }

