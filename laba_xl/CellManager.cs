using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace laba_xl
{
    class CellManager
    {
        private DataGridView _dataGridView;
        private static CellManager _instance;

        public Cell CurrentCell { get; set; }
        public DataGridView DataGridView { set { _dataGridView = value; } }

        // Lets you keep the one and only CellManager instance
        // during the whole runtime, according to Singleton pattern.
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

        // Returns a cell by its string-represented coordinates in the data grid view.
        public Cell GetCell(string cellName)
        {
            var matches = new Regex(@"^R(?<row>\d+)C(?<col>\d+)$").Matches(cellName);
            int row = Int32.Parse(matches[0].Groups["row"].Value) - 1;
            int col = Int32.Parse(matches[0].Groups["col"].Value) - 1;

            try
            {
                return (Cell)_dataGridView[col, row].Tag;
            }
            catch (ArgumentOutOfRangeException)
            {
                return null;
            }
        }

        // Returns a cell by its actual representative parent in the data grid view.
        public Cell GetCell(DataGridViewCell dataGridViewCell)
        {
            return (Cell)dataGridViewCell.Tag;
        }

        // Checks if a cell contains the reference to itself
        // (probably throughout the chain of several nested references).
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

            return HasInnerRecursion(cell, invokerName);
        }

        private bool HasInnerRecursion(Cell cell, string invokerName)
        {
            List<Cell> refs = cell.CellReferences;

            for (int i = refs.Count - 1; i >= 0; i--)
            {
                if (refs[i].Name.Equals(""))
                {
                    return false;
                }

                if (refs[i].Name.Equals(CurrentCell.Name) || HasReferenceRecursion(refs[i], invokerName))
                {
                    return true;
                }
            }

            return false;
        }
    }
}
