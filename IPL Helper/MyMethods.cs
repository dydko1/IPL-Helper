using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Diagnostics;

namespace IPL_Helper
{
    public partial class Form1
    {
        #region GetSheets
        //get data from Excel
        private bool GetSheet()
        {
            Excel.Range rangeFind = null;
            //connectToExcel();
            if (!ConnectToExcel())
            {
                //MessageBox.Show("Run Excel. Active instance doesn't exist", this.Name, MessageBoxButtons.OKCancel, MessageBoxIcon.Stop);
                return false;
            }

            //int start=0, stop=0;
            //start = Environment.TickCount & Int32.MaxValue;
            try
            {
                //  range = xlWorkSheet.UsedRange; 
                lastRowSheet = xlWorkSheet.Range["C" + xlWorkSheet.Rows.Count].End[Excel.XlDirection.xlUp].Row; // Return last cell in column
                range = xlWorkSheet.get_Range("C1", "C" + lastRowSheet.ToString());
                //string ss = range.Address.ToString().Substring(range.Address.ToString().LastIndexOf("$")+1) ; //.LastIndexOf("$").ToString(); // cut last $

                // You should specify all these parameters every time you call this method,
                // since they can be overridden in the user interface. 
                rangeFind = range.Find("DESCRIPTION", Missing.Value, Excel.XlFindLookIn.xlValues,
                        Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows,
                        Excel.XlSearchDirection.xlNext, false,
                        Missing.Value, Missing.Value);

                if (rangeFind == null) //if null abort
                {
                    MessageBox.Show("Title text 'DESCRIPTION' not found in cells.\nLook on Excel- column $C.\nPlease Excel data check again.", "Active sheet check");
                    //if(currentWorkSheet == null)
                    //currentWorkSheet = null;
                    this.Text = "Select current sheet !!";
                    return false;
                }

                // get range
                firstRowSheet = rangeFind.Row + 1;
                range = xlWorkSheet.get_Range("B" + (firstRowSheet), "C" + lastRowSheet); // Set new range   

                FillDataGrid();
                // !!!!
                currentWorkSheet = xlWorkSheet;
                this.Text = "Current sheet is: " + currentWorkSheet.Name;

                // !!!!
            }
            catch (Exception ex)
            {
                xlApp = null;
                xlWorkSheet = null;
                rangeFind = null;
                range = null;
                MessageBox.Show("Excel Check data.\n " + ex.ToString(), this.Name, MessageBoxButtons.OKCancel, MessageBoxIcon.Stop);
            }
            return true;
        }

        // check Excel is running
        private bool ConnectToExcel()
        {
            // Check Excel run processes
            Process[] excelProcess = Process.GetProcessesByName("EXCEL");
            if (excelProcess.Length == 0)
            {
                MessageBox.Show("Run Excel. Active instance doesn't exist !\n", this.Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            else if (excelProcess.Length >= 2)
            {
                MessageBox.Show("Too many EXCEL process are active. \nExcel shut down (kill) and look in TASK MANAGER active process.\n"
                    + excelProcess.Length.ToString() + " Excel processes are active. \nPlease leave one active instances.", this.Name, MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            }

            // Try connect to active instance Excel
            try
            {
                xlApp = (Excel.Application)System.Runtime.InteropServices.Marshal.
                    GetActiveObject("Excel.Application");
                xlWorkSheet = xlApp.ActiveSheet;
                if (xlWorkSheet == null)
                    throw new NullReferenceException();
                //  xlApp.Visible = false;
                return true;
            }
            catch (NullReferenceException ex)  // I think process run in backgroud
            {
                MessageBox.Show("Some Excel process run in backgroud. Run Tak Manager and check.\n" + ex.ToString(), this.Name, MessageBoxButtons.OKCancel, MessageBoxIcon.Stop);
                xlApp = null;
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Run Excel. Active instance doesn't exist !\n" + ex.ToString(), this.Name, MessageBoxButtons.OKCancel, MessageBoxIcon.Stop);
                xlApp = null;
                return false;
            }
        }

        // Datagrid fill data from range
        private void FillDataGrid()
        {
            Array tempRange = null;
            try
            {
                //MessageBox.Show(firstRowSheet.ToString() + "\t" +lastRowSheet.ToString());
                tempRange = (Array)range.Cells.Value2; // add ToString() throw exception 
                int rowCount = tempRange.GetLength(0) + 1; // length range
                dataGridView1.Rows.Clear();
                // Import/copy data to datagrid from Array
                for (int rowIndex = 1; rowIndex < rowCount; ++rowIndex)
                {
                    if (tempRange.GetValue(rowIndex, 1) != null && tempRange.GetValue(rowIndex, 2) != null)
                        dataGridView1.Rows.Add(tempRange.GetValue(rowIndex, 1).ToString(), tempRange.GetValue(rowIndex, 2).ToString());
                    else if (tempRange.GetValue(rowIndex, 1) == null && tempRange.GetValue(rowIndex, 2) != null)
                        dataGridView1.Rows.Add("", tempRange.GetValue(rowIndex, 2).ToString());
                    else
                    {
                        MessageBox.Show("Cell in Excel Colum $C is empty! Remove empty row/rows.", "Excel check data in $C cell!", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop);
                        return;
                    }
                }
                SetRowNumber();
                SetLevels();// Delete after used;
            }
            catch (Exception ex)
            {
                xlApp = null;
                xlWorkSheet = null;
                range = null;
                tempRange = null;
                MessageBox.Show("Excel check data.\n " + ex.ToString(), this.Name, MessageBoxButtons.OKCancel, MessageBoxIcon.Stop);
            }
            SetColors();
        }
        
        // ustwia numery wierszy
        private void SetRowNumber()
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                row.HeaderCell.Value = (row.Index + firstRowSheet).ToString();
                //row.HeaderCell.Value = (row.Index + 0).ToString();
            }
            //dataGridView1.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }

        // ustawia levele w 2 kolumnie
        private void SetLevels()
        {
            int countRows = dataGridView1.Rows.Count - 1;

            if (countRows <= 1)
                MessageBox.Show("BOM is to short!", "BOM is to short!");

            for (int i = 0; i <= countRows; i++)
                dataGridView1.Rows[i].Cells[2].Value = CountDot(dataGridView1.Rows[i].Cells[0].Value.ToString());
        }

        public static int CountDot(string text)
        {
            char pattern = '.'; //separator //kiedys do poprawy
            int count = 0; //jak jest ...... to zle zlicza z datagrid

            if (string.IsNullOrEmpty(text))
            {
                MessageBox.Show("Cell in column $B sheet is empty", "Empty cell", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 0;
                //throw new System.ArgumentException("Parameter cannot be null or empty", "Method: CountDot");
            }
            foreach (char ch in text)   // jak jest napis text ="..." to jest ok ?? moze kiedys sie porpawi
            {
                if (ch == pattern)
                    count++;
            }
            return count; //return occurences of "."
        }

        private void SetColors()
        {
            int countRows = dataGridView1.Rows.Count - 1, level = -1;
            /***********/
            //DataGridViewCellStyle normalFont = new DataGridViewCellStyle();
            DataGridViewCellStyle boldFont = new DataGridViewCellStyle();
            boldFont.Font = new Font(Font.FontFamily, Font.Size, FontStyle.Bold);
            /***********/
            for (int i = 0; i <= countRows; i++)
            {
                level = (int)dataGridView1.Rows[i].Cells[2].Value;
                //^^^^^^^^^^^^
                if (level == 0 || level > 4) //domyślne kolory jak 0 się ładują
                {
                    if (countRows == i)//top
                        dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.FromArgb(217, 217, 217);
                }
                else if (level == 1)
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.FromArgb(216, 228, 188);
                }
                else if (level == 2)
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.FromArgb(196, 215, 155);
                }
                else if (level == 3)
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.FromArgb(118, 147, 60);
                }
                else if (level == 4)
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.FromArgb(79, 98, 40);
                }
                else
                {
                    MessageBox.Show("Incorrect value", "Incorrect value");
                }

                if (IsBOM(i) || (countRows == i)) //jest BOM lub TOP Level
                    dataGridView1.Rows[i].DefaultCellStyle.Font = boldFont.Font;
                //^^^^^^^^^^^^
                //dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.BlueViolet; //Color.FromArgb(0, 0, 0);
                //dataGridView1.Rows[i].DefaultCellStyle.Font = boldFont.Font;
                //dataGridView1.Rows[i].DefaultCellStyle.Font = normalFont.Font;
            }
        }
        #endregion

        #region Move Down Up
        // check Excel current sheet
        private bool CheckActiveSheet()
        {
            try
            {
                if (xlApp.ActiveSheet.Name == xlWorkSheet.Name)
                {
                    return true;
                }
                else
                {
                    DialogResult dialogResult = MessageBox.Show
                        (" Excel sheet selecting and import data!\n Last data imported from sheet:\t "
                            + xlWorkSheet.Name + "\n Do you want import/download NEW data from Excel active sheet Y/n?: \t ",
                            "Be careful when you select sheet !", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dialogResult == DialogResult.Yes)
                    {
                        if (!GetSheet()) //get all data
                        {
                            return false;
                        }
                    }
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Run Excel. Excel active workbook doesn't exist !\n" + ex.ToString(), this.Name, MessageBoxButtons.OKCancel, MessageBoxIcon.Stop);
                return false;
            }
        }

        private bool GetMinMaxLevel(ref int selectFirstRow, ref int selectLastRow, ref int selectedLevelDot, int countRows)
        {
            if (countRows >= 2)
            {
                for (int i = 0; i <= countRows - 1; i++)
                {
                    if (selectFirstRow == -1)
                    {
                        if (dataGridView1.Rows[i].Selected)
                        {
                            selectFirstRow = i;
                            selectLastRow = i;
                        }
                    }
                    if (selectLastRow >= 0)
                    {
                        if (dataGridView1.Rows[i].Selected)
                        {
                            selectLastRow = i;
                        }
                    }
                }
                if (selectFirstRow == -1 || selectLastRow == -1)
                {
                    MessageBox.Show("Select row/rows", "Select row/rows");
                    return false;
                }
                else
                {
                    selectedLevelDot = (int)dataGridView1.Rows[selectLastRow].Cells[2].Value;
                    return true;
                }
            }
            else
            {
                MessageBox.Show("BOM is to short!", "BOM is to short!");
                return false;
            }
        }

        private bool IsCurrentSelect()
        {
            if (selectRange >= 1 && (selectRange == dataGridView1.SelectedRows.Count))
                return true;
            else
                return false;
        }

        private void MoveUp(int selectFirstRow, int selectLastRow, int selectedLevelDot, int countRows)
        {
            //address occurrences-  selectFirstRow -> example 150.2 & selectLastRow -> example 150 -> selectLastRow > selectFirstRow
            //occurrences address, selectedLevelDot Dot count occurences, countRows -1 as Array,
            //checkNextRow next new position
            bool isPrevBOM = false;
            //IsBOM(selectFirstRow);
            int countPreviusLevelDot = -1, checkPreviusRow = -1, endPreviusRow = -1;
            string selectAddr = "", insertAddr = "", sheetRange = range.Address.ToString(); // sheetRange traci zakres ggy wiersz jest przenoszony na pierwszą pozycję
            Excel.Range rng = null;

            if (selectFirstRow >= 2)
                isPrevBOM = IsBOM(selectFirstRow - 1);

            for (int i = selectFirstRow - 1; i >= 0; i--) // BOM --> -1 +  && i + 1 --> -2
            {
                checkPreviusRow = i;
                countPreviusLevelDot = (int)dataGridView1.Rows[checkPreviusRow].Cells[2].Value;

                if (selectedLevelDot > countPreviusLevelDot)
                {
                    //MessageBox.Show("papapa\t" + checkPreviusRow.ToString());
                    return;
                }

                if (isPrevBOM && (checkPreviusRow >= 1))
                {
                    countPreviusLevelDot = (int)dataGridView1.Rows[checkPreviusRow - 1].Cells[2].Value;
                    if (selectedLevelDot == countPreviusLevelDot)
                    {
                        //checkPreviusRow--;
                        break;
                    }
                }
                else
                {
                    //MessageBox.Show(checkPreviusRow.ToString());
                    break;
                }
            }

            if (checkPreviusRow != -1) //&& (selectedLevelDot == countPreviusLevelDot))
            {
                endPreviusRow = checkPreviusRow + selectLastRow - selectFirstRow;
                try
                {
                    selectAddr = Convert.ToString(firstRowSheet + selectFirstRow) + ":" +
                        Convert.ToString(firstRowSheet + selectLastRow);
                    insertAddr = Convert.ToString(firstRowSheet + checkPreviusRow) + ":" +
                        Convert.ToString(firstRowSheet + checkPreviusRow);
                    //    Convert.ToString(firstRowSheet + endPreviusRow); // stara obliczona pozycja wychodzi na to samo
                    // MessageBox.Show(selectAddr + "\t" + insertAddr + "\t" + (firstRowSheet + endPreviusRow));

                    rng = xlWorkSheet.get_Range(selectAddr); //(, Missing.Value)
                    rng.EntireRow.Select();
                    rng.EntireRow.Cut();
                    rng = xlWorkSheet.get_Range(insertAddr); //(, Missing.Value)
                    rng.EntireRow.Insert(); //(Excel.XlDirection.xlDown); need ??
                    //xlApp.CutCopyMode = (Excel.XlCutCopyMode)0;
                    range = xlWorkSheet.get_Range(sheetRange);
                    // http://stackoverflow.com/questions/24610344/better-way-to-insert-excel-range-as-row-and-copy-formats-using-c-sharp
                    // Rows("48:48").Select
                    //'Range("A65").Activate
                    //Selection.Cut
                    //Rows("59").Select 'nowa pozycja
                    // Selection.Insert Shift:=xlUp 'lub xlDown
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Move Down Error!" + ex.ToString(), "Move Down!");
                }

                FillDataGrid(); //nie moze byc bo problem
                //GetSheet(); //tu problem 
                dataGridView1.CurrentCell = dataGridView1[0, endPreviusRow];
                //dataGridView1.Rows[checkPreviusRow + selectLastRow - selectFirstRow].Frozen = true;
                SelectRowsBOM(endPreviusRow);
            }
            else
            {
                //MessageBox.Show("Unexpected error Move down");
                return;
            }
            /*  http://stackoverflow.com/questions/1012708/datagridview-selected-row-move-up-and-down
            dataGridView1.Rows[5].DefaultCellStyle.BackColor = Color.Red;
            var rows = dataGridView1.Rows;
            var prevRow = rows[3];
            rows.Remove(prevRow);
            prevRow.Frozen = false;
            rows.Insert(5, prevRow);
            dataGridView1.ClearSelection();
            dataGridView1.Rows[5 - 1].Selected = true;
            */
        }

        private void MoveDown(int selectFirstRow, int selectLastRow, int selectedLevelDot, int countRows)
        {
            //address occurrences-  selectFirstRow -> example 150.2 & selectLastRow -> example 150 -> selectLastRow > selectFirstRow
            //occurrences address, selectedLevelDot Dot count occurences, countRows -1 as Array,
            //checkNextRow next new position
            int countNextLevelDot = -1, checkNextRow = -1; //, endPreviusRow = -1;
            string selectAddr = "", insertAddr = "", sheetRange = range.Address.ToString();
            Excel.Range rng = null;
            for (int i = selectLastRow; i <= countRows - 2; i++) // BOM --> -1 +  && i + 1 --> -2
            {
                checkNextRow = i + 1;
                countNextLevelDot = (int)dataGridView1.Rows[checkNextRow].Cells[2].Value;
                if (selectedLevelDot == countNextLevelDot)
                {
                    break;
                }
                else if (selectedLevelDot > countNextLevelDot)
                    return;
            }

            if ((checkNextRow != -1) && (selectedLevelDot == countNextLevelDot))
            {

                //endPreviusRow = checkNextRow + selectLastRow - selectFirstRow;
                try
                {
                    selectAddr = Convert.ToString(firstRowSheet + selectFirstRow) + ":" +
                        Convert.ToString(firstRowSheet + selectLastRow);
                    insertAddr = Convert.ToString(firstRowSheet + checkNextRow + 1) + ":" +
                        Convert.ToString(firstRowSheet + checkNextRow + 1);
                    //  Convert.ToString(firstRowSheet + checkNextRow + selectLastRow - selectFirstRow + 1); // stara obliczona pozycja wychodzi na to samo 
                    //  MessageBox.Show(selectAddr + "\t" + insertAddr + "\t" + checkNextRow);

                    rng = xlWorkSheet.get_Range(selectAddr); //(, Missing.Value)
                    rng.EntireRow.Select();
                    rng.EntireRow.Cut();
                    rng = xlWorkSheet.get_Range(insertAddr); //(, Missing.Value)
                    rng.EntireRow.Insert(); //(Excel.XlDirection.xlDown); need ?? to zajmuje duzo czasu
                    range = xlWorkSheet.get_Range(sheetRange);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Move Down Error!" + ex.ToString(), "Move Down!");
                }

                FillDataGrid(); //nie moze byc bo problem
                dataGridView1.CurrentCell = dataGridView1[0, checkNextRow];
                SelectRowsBOM(checkNextRow);
            }
            else
            {
                //MessageBox.Show("Unexpected error Move down");
                return;
            }
        }

        private bool IsBOM(int curentRow)
        {
            if (curentRow < 1)
            {
                return false;
            }
            else if ((int)dataGridView1.Rows[curentRow].Cells[2].Value < (int)dataGridView1.Rows[curentRow - 1].Cells[2].Value)
            {
                return true;
            }
            else
                return false;
        }

        private void SelectRowsBOM(int clickCell)
        {
            // last rowCount is TopAssy but without Assy = -2, first row = 0
            // address rowClick 0...n-1,  rowCount 1...n
            int checkCell = clickCell, rowCount = dataGridView1.Rows.Count - 1, countClickDot = 0;

            if (checkCell >= 0 && checkCell < rowCount) // Need "checkCell >= 0 &&" when click DataGrid headline name.
            {
                int previousDots = 0; // back cell
                dataGridView1.ClearSelection(); // clear datagrid
                dataGridView1.Rows[checkCell].Selected = true;
                countClickDot = (int)dataGridView1.Rows[checkCell].Cells[2].Value;

                while (checkCell >= 1)
                {
                    checkCell--;
                    previousDots = (int)dataGridView1.Rows[checkCell].Cells[2].Value;
                    if (countClickDot < previousDots)
                    {
                        dataGridView1.Rows[checkCell].Selected = true;
                    }
                    else
                    {
                        break;
                    }
                }
                selectRange = dataGridView1.SelectedRows.Count;
            }
            else if (checkCell == rowCount)
            {
                dataGridView1.ClearSelection();
                MessageBox.Show("Don't click/move Top ASSY !", "Don't click/move Top ASSY !");
            }
            else
            {
                dataGridView1.ClearSelection();
                MessageBox.Show("You must click in row !", "You must click in row !");
            }
        }

        #endregion

        #region Delete
        void xlApp_WorkbookBeforeClose(Excel.Workbook MyWorkbook, ref bool Cancel)
        {
            // W innej kombinacji Excel sie wysypuje, Allert Crash, Bug in Excel
            if (MyWorkbook.FullName != temPath)
            {
                Cancel = false;
            }
            else
            {
                Cancel = true;
                CloseTemplate();
            }
        }

        private void CloseTemplate()
        {
            if (xlTemWorkBook != null || xlTemSheet != null)
            {
                //MessageBox.Show("Close the template: " + xlTemWorkBook.Name);
                xlApp.WorkbookBeforeClose -= new Excel.AppEvents_WorkbookBeforeCloseEventHandler(xlApp_WorkbookBeforeClose);
                xlApp.CutCopyMode = (Excel.XlCutCopyMode)0;
                //xlApp.DisplayAlerts = false;
                xlTemWorkBook.Close(false, Type.Missing, Type.Missing);
                xlApp.CutCopyMode = (Excel.XlCutCopyMode)1;
                //xlApp.DisplayAlerts = true;
                xlTemSheet = null;
                xlTemWorkBook = null;
            }
        }

        #endregion

        #region Rename
        private int IndexOfDots(string s, int level)
        {
            int dotPosition = -1, i = 1;

            if (string.IsNullOrEmpty(s))
            {
                MessageBox.Show("Null or Empty cell !", "Null or Empty cell !");
                return -1;
            }

            while (i <= level && (dotPosition = s.IndexOf(".", dotPosition + 1)) != -1) // ladnie dziala 
            {
                if (level == i)
                    return dotPosition;
                i++;
            }
            return -1;
        }
        #endregion
    }
}
