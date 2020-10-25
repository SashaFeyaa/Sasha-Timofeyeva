using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LabCalculator
{
    public partial class Form1 : Form
    {
        private const int _maxCols = 10;
        private const int _maxRows = 10;
        private const int _rowHeaderWidth = 100;
        private const int _colHeaderWidth = 30;
        private bool _formulaView = false;
        private bool _errorOccured = false;
        private const string ERROR_LAST_ROW = "You cant delete the last row";
        private const string ERROR_LAST_COL = "You cant delete the last col";
        private const string ERROR_DEPENDENCIES_ROW = "There are references in the table to row cells, which you are trying delete. Please, delete them and try again";
        private const string ERROR_DEPENDENCIES_COLUMN = "There are references in the table to column cells, which you are trying delete. Please, delete them and try again";
        private const string CAPTION_DELETE_ROW = "Row deleting";
        private const string CAPTION_DELETE_COL = "Column deleting";
        private const string CAPTION_ERR = "ERROR";
        private const string CAPTION_WARNING = "WARNING";
        private const string WARNING_ERROR = "Correct mistake in the table before saving";
        private string _currentFilePath = "";
        private const string CAPTION_BASICS = "General information";
        private const string CAPTION_ERRORS = "Errors";
        private const string CAPTION_FEATURES = "Special features";
        private const string CAPTION_CREDITS = "Author";
        private const string FPATH_LOOKUP = "C:\\Users\\Asus\\Desktop\\lab2\\lookup.txt";
        private const string FPATH_ERRORS = "C:\\Users\\Asus\\Desktop\\lab2\\errors.txt";
        private const string FPATH_FEATURES = "C:\\Users\\Asus\\Desktop\\lab2\\features.txt";
        private const string FPATH_CREDITS = "C:Users\\Asus\\Desktop\\lab2\\credits.txt";
        private const string WARNING_SAVE = "Save the current table?";
        private const string ASK_DELETE_ROW = "DELETE ROW?";
        private const string ASK_DELETE_COL = "DELETE COL?";
        private const string ERROR_DEPENDENCIES_COL = "ERROR";
       

        private void InitializeDataGridView()
        {
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.ColumnCount = _maxCols;
            dataGridView1.RowCount = _maxRows;


            FillHeaders();

            dataGridView1.AutoResizeRows();
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.RowHeadersWidth = _rowHeaderWidth;
            dataGridView1.ColumnHeadersHeight = _colHeaderWidth;
        }

        private void FillHeaders()
        {
            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
                col.HeaderText = "C" + (col.Index + 1);
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                row.HeaderCell.Value = "R" + (row.Index + 1);
            }
        }

        private void InitializeAllCells()
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                //  row.DefaultCellStyle.Font = new Font(_defaultFontName, _defaultFontSize, GraphicsUnit.Point);

                foreach (DataGridViewCell cell in row.Cells)
                {
                    InitializeSingleCell(row, cell);
                }

            }
        }

        private void InitializeSingleCell(DataGridViewRow row, DataGridViewCell cell)
        {
            string cellName = "R" + (row.Index + 1) + "C" + (cell.ColumnIndex + 1).ToString();
            cell.Tag = new Cell(cell, cellName, "0");
            cell.Value = "0";
        }


        private void UpdateCellValues()
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                foreach (DataGridViewCell dgvCell in row.Cells)
                {
                    UpdateSingleCellValue(dgvCell);
                }
            }
        }


        private void UpdateCellValues(DataGridViewCell invoker)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                foreach (DataGridViewCell dgvCell in row.Cells)
                {
                    if (invoker != dgvCell)
                    {
                        UpdateSingleCellValue(dgvCell);
                    }
                }
            }
        }


        private void UpdateSingleCellValue(DataGridViewCell dgvCell)
        {
            Cell cell = (Cell)dgvCell.Tag;

            if (!_formulaView)
            {
                if (cell.Formula.Equals("") || Regex.IsMatch(cell.Formula, @"^\d+$"))
                {
                    dgvCell.Value = cell.Value;
                }
                else
                {
                    dgvCell.Value = cell.Evaluate();
                }
            }
            else
            {
                dgvCell.Value = cell.Formula;
            }
        }

        public Form1()
        {
            InitializeComponent();
            WindowState = FormWindowState.Maximized;

            InitializeDataGridView();
            InitializeAllCells();
            CellManager.Instance.DataGridView = dataGridView1;

        }



        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            if (e.RowIndex == -1 || e.ColumnIndex == -1)
            {
                return;
            }

            Cell cell = (Cell)dataGridView1[e.ColumnIndex, e.RowIndex].Tag;
            DataGridViewCell dgvCell = cell.Parent;

            if (!dgvCell.ReadOnly)
            {
                dataGridView1.BeginEdit(true);
            }

        }





        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            Cell cell = (Cell)dataGridView1[e.ColumnIndex, e.RowIndex].Tag;

            CellManager.Instance.CurrentCell = cell;
            DataGridViewCell dgvCell = cell.Parent;
            dgvCell.Value = cell.Formula;

        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

            Cell cell = (Cell)dataGridView1[e.ColumnIndex, e.RowIndex].Tag;
            DataGridViewCell dgvCell = cell.Parent;

            if (dgvCell.Value == null)
            {
                cell.Formula = "0";
                cell.Value = "0";
                dgvCell.Value = "0";
            }

            ClearRemovedReferences(cell);
            ResolveCellFormula(cell, dgvCell);
        }


        private void SaveDataGridView(string filePath)
        {
            _currentFilePath = filePath;
            dataGridView1.EndEdit();

            DataTable table = new DataTable("data");
            ForgetDataTable(table);
            table.WriteXml(filePath);
        }


        private void ForgetDataTable(DataTable table)
        {
            foreach (DataGridViewColumn dgvColumn in dataGridView1.Columns)
            {
                table.Columns.Add(dgvColumn.Index.ToString());
            }

            foreach (DataGridViewRow dgvRow in dataGridView1.Rows)
            {
                DataRow dataRow = table.NewRow();

                foreach (DataColumn col in table.Columns)
                {
                    Cell cell = (Cell)dgvRow.Cells[Int32.Parse(col.ColumnName)].Tag;
                    dataRow[col.ColumnName] = cell.Formula;
                }

                table.Rows.Add(dataRow);
            }
        }


        private bool SaveDataGridView(string filePath, string dummy)
        {
            if (!filePath.Equals(""))
            {
                SaveDataGridView(filePath);
                return true;
            }
            else if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                SaveDataGridView(saveFileDialog1.FileName);
                return true;
            }
            return false;
        }


        private void LoadDataGridView(string filePath)
        {
            _currentFilePath = filePath;
            DataSet dataSet = new DataSet();
            dataSet.ReadXml(filePath);
            DataTable table = dataSet.Tables[0];

            dataGridView1.ColumnCount = table.Columns.Count;
            dataGridView1.RowCount = table.Rows.Count;

            foreach (DataGridViewRow dgvRow in dataGridView1.Rows)
            {
                foreach (DataGridViewCell dgvCell in dgvRow.Cells)
                {
                    string cellName = "R" + (dgvRow.Index + 1).ToString() + "C" + (dgvCell.ColumnIndex + 1).ToString();
                    string formula = table.Rows[dgvCell.RowIndex][dgvCell.ColumnIndex].ToString();
                    dgvCell.Tag = new Cell(dgvCell, cellName, formula);
                }
            }

            UpdateCellValues();
        }


        private void ClearRemovedReferences(Cell cell)
        {
            List<Cell> removedCells = new List<Cell>();

            foreach (Cell refCell in cell.CellReferences)
            {
                if (!cell.Formula.Contains(refCell.Name))
                {
                    removedCells.Add(refCell);
                }
            }
        }

        private void ResolveCellFormula(Cell cell, DataGridViewCell dgvCell)
        {
            cell.Formula = dgvCell.Value.ToString();
            string cellValue = cell.Evaluate().ToString();

            if (!cell.Error.Equals(""))
            {
                _errorOccured = true;
                MessageBox.Show(cell.Error, CAPTION_ERR, MessageBoxButtons.OK, MessageBoxIcon.Error);
                cell.Error = "";
                DisableCellsButCurrent(dgvCell);
            }
            else
            {
                _errorOccured = false;
                dgvCell.Value = _formulaView ? cell.Formula : cellValue;
                EnableCells();
                UpdateCellValues(dgvCell);
            }
        }

        private void DisableCellsButCurrent(DataGridViewCell current)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    cell.ReadOnly = true;
                    cell.Style.BackColor = Color.LightSalmon;
                    cell.Style.ForeColor = Color.DarkSalmon;
                }
            }

            current.ReadOnly = false;
            current.Style.BackColor = current.OwningColumn.DefaultCellStyle.BackColor;
            current.Style.ForeColor = current.OwningColumn.DefaultCellStyle.ForeColor;

        }

        private void EnableCells()
        {

            if (!_errorOccured)
            {
                return;
            }

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    cell.ReadOnly = false;
                    cell.Style.BackColor = cell.OwningColumn.DefaultCellStyle.BackColor;
                    cell.Style.ForeColor = cell.OwningColumn.DefaultCellStyle.ForeColor;
                }
            }

        }

        private void DataGridView_CellStateChanged(object sender, DataGridViewCellStateChangedEventArgs e)
        {
            if (e.Cell.ReadOnly)
            {
                e.Cell.Selected = false;
            }
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                CellManager.Instance.CurrentCell = new Cell();
                LoadDataGridView(openFileDialog1.FileName);
            }
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            if (_errorOccured)
            {
                MessageBox.Show(WARNING_ERROR, CAPTION_WARNING, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SaveDataGridView(_currentFilePath, "");
        }

        private void AddRow()
        {
            dataGridView1.Rows.Add(new DataGridViewRow());
            FillHeaders();

            DataGridViewRow addedRow = dataGridView1.Rows[dataGridView1.RowCount - 1];
            // addedRow.DefaultCellStyle.Font = new Font(_defaultFontName, _defaultFontSize, GraphicsUnit.Point);

            foreach (DataGridViewCell cell in addedRow.Cells)
            {
                InitializeSingleCell(addedRow, cell);
            }
        }

        private void AddColumn()
        {
            dataGridView1.Columns.Add(new DataGridViewColumn(dataGridView1.Rows[0].Cells[0]));
            FillHeaders();

            foreach (DataGridViewRow dgvRow in dataGridView1.Rows)
            {
                InitializeSingleCell(dgvRow, dgvRow.Cells[dataGridView1.ColumnCount - 1]);

            }
        }

        private void DeleteRow()
        {
            DialogResult result = MessageBox.Show(ASK_DELETE_ROW, CAPTION_DELETE_ROW, MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                if (dataGridView1.RowCount == 1)
                {
                    MessageBox.Show(ERROR_LAST_ROW, CAPTION_ERR, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (DeletedRowHasDependencies())
                {
                    MessageBox.Show(ERROR_DEPENDENCIES_ROW, CAPTION_ERR, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                int lastRowInd = dataGridView1.RowCount - 1;
                dataGridView1.Rows.RemoveAt(lastRowInd);

            }
        }

        private bool DeletedRowHasDependencies()
        {
            List<string> deletedNames = new List<string>();
            int lastInd = dataGridView1.RowCount - 1;

            foreach (DataGridViewCell dgvCell in dataGridView1.Rows[lastInd].Cells)
            {
                Cell cell = (Cell)dgvCell.Tag;
                deletedNames.Add(cell.Name);
            }
            return FindDeletedRowDependenciesInTable(deletedNames, lastInd);
        }

        private bool FindDeletedRowDependenciesInTable(List<string> deletedNames, int lastInd)
        {
            for (int i = 0; i < lastInd; i++)
            {
                foreach (DataGridViewCell dgvCell in dataGridView1.Rows[i].Cells)
                {
                    Cell cell = (Cell)dgvCell.Tag;
                    List<Cell> refs = cell.CellReferences;

                    for (int j = refs.Count - 1; j >= 0; j--)
                    {
                        if (deletedNames.Contains(refs[j].Name))
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        private void DeleteColumn()
        {
            DialogResult result = MessageBox.Show(ASK_DELETE_COL, CAPTION_DELETE_COL, MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                if (dataGridView1.ColumnCount == 1)
                {
                    MessageBox.Show(ERROR_LAST_COL, CAPTION_ERR, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (DeletedColumnHasDependencies())
                {
                    MessageBox.Show(ERROR_DEPENDENCIES_COL, CAPTION_ERR, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                int lastColInd = dataGridView1.ColumnCount - 1;
             
                dataGridView1.Columns.RemoveAt(lastColInd);

            }
        }


        private bool DeletedColumnHasDependencies()
        {
            List<string> deletedNames = new List<string>();
            int lastInd = dataGridView1.ColumnCount - 1;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                
                Cell cell = (Cell)row.Cells[lastInd].Tag;
                deletedNames.Add(cell.Name);
            }
            return FindDeletedColumnDependenciesInTable(deletedNames, lastInd);
        }

        private bool FindDeletedColumnDependenciesInTable(List<string> deletedNames, int lastInd)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                for (int i = 0; i < lastInd; i++)
                {
                    Cell cell = (Cell)row.Cells[i].Tag;
                    List<Cell> refs = cell.CellReferences;

                    for (int j = refs.Count - 1; j >= 0; j--)
                    {
                        if (deletedNames.Contains(refs[j].Name))
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        private void AddRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddRow();
        }

        private void AddColumnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddColumn();
        }

        private void DeleteRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DeleteRow();
        }

        private void DeleteColumnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DeleteColumn();
        }

        private void ValesToolStripMenuItem_Click(object sender, EventArgs e)
        {
         /*
            valuesToolStripeMenuItem.Checked = true;
            valuesToolStripeMenuItem.Enabled = false;
            formulasToolStripMenuItem.Checked = false;
            formulasToolStripMenuItem.Enabled = true;*/
            ToggleFormulaView();
        }

        private void FormulasToolStripMenuItem_Click(object sender, EventArgs e)
        {
          /*  valuesToolStripeMenuItem.Checked = true;
            valuesToolStripeMenuItem.Enabled = false;
            formulasToolStripMenuItem.Checked = false;
            formulasToolStripMenuItem.Enabled = true;*/
            ToggleFormulaView();
        }

        private void ToggleFormulaView()
        {
            _formulaView = !_formulaView;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                foreach (DataGridViewCell dgvCell in row.Cells)
                {
                    Cell cell = CellManager.Instance.GetCell(dgvCell);

                    if (_formulaView)
                    {
                        dgvCell.Value = cell.Formula;
                    }
                    else
                    {
                        dgvCell.Value = cell.Value;
                    }
                }
            }
        }


        private void LookupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string helpText = System.IO.File.ReadAllText(FPATH_LOOKUP);
            MessageBox.Show(helpText, CAPTION_BASICS, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void FeaturesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string helpText = System.IO.File.ReadAllText(FPATH_FEATURES);
            MessageBox.Show(helpText, CAPTION_FEATURES, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void CreditsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string helpText = System.IO.File.ReadAllText(FPATH_CREDITS);
            MessageBox.Show(helpText, CAPTION_CREDITS, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Tabledit_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult result = MessageBox.Show(WARNING_SAVE, CAPTION_WARNING, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                if (_errorOccured)
                {
                    MessageBox.Show(WARNING_ERROR, CAPTION_WARNING, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }

                if (!SaveDataGridView(_currentFilePath, ""))
                {
                    e.Cancel = true;
                }
            }
            else if (result == DialogResult.Cancel)
            {
                e.Cancel = true;
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}

    

       