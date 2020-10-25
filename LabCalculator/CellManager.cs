using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace LabCalculator
{
    class CellManager
    {

        private DataGridView _dataGridView;
        private static CellManager _instance;

        public Cell CurrentCell { get; set; }

        public DataGridView DataGridView { set { _dataGridView = value; } }


        public static CellManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new CellManager();
                }
                return _instance;
            }
        }
            private CellManager() { }


        public Cell GetCell(string cellName)
        {
            var matches = new Regex(@"^R(?<row>\d+)C(?<col>\d+)$").Matches(cellName);
            int row = Int32.Parse(matches[0].Groups["row"].Value) - 1;
            int col = Int32.Parse(matches[0].Groups["col"].Value) - 1;

            try
            {
                return (Cell)_dataGridView[col, row].Tag;
            }
            catch(ArgumentOutOfRangeException)
            {
                return null;
            }

        }

        public Cell GetCell(DataGridViewCell dataGridViewCell)
        {
            return (Cell)dataGridViewCell.Tag;
        }


        public bool HasReferenceRecursion(Cell cell, string invokerName)
        {
            string cellName = cell.Name;

            if (cellName.Equals("") || !invokerName.Equals(CurrentCell.Name))
            {
                return false;
            }

            if (cellName.Equals(CurrentCell.Name))
            {
                return true;
            }

            return HasReferenceRecursion(cell, invokerName);
        }

        private bool HasInnerRecursion(Cell cell, string invokerName)
        {
            List<Cell> refs = cell.CellReferences;

            for(int i = refs.Count - 1; i >= 0; i--)
            {
                if (refs[i].Name.Equals(""))
                {
                    return false;
                }

                if(refs[i].Name.Equals(CurrentCell.Name) || HasReferenceRecursion(refs[i], invokerName))
                {
                    return true;
                }
            }

            return false;
        }

    }
}
