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
using System.IO;
using IPL_Helper.Properties;

namespace IPL_Helper
{
    public partial class Form1 : Form
    {
        Excel.Application xlApp = null;
        Excel.Workbook xlTemWorkBook = null;
        Excel.Worksheet xlWorkSheet = null, xlTemSheet = null, currentWorkSheet = null;
        Excel.Range range = null;
        string temPath = "";
        int firstRowSheet = 0, lastRowSheet = 0, selectRange = -1; // Excel range first and last cell 
        
        public Form1()
        {
            InitializeComponent();
            if (!GetSheet())
            {
                return;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!GetSheet())
            {
                return;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //address occurrences-  selectFirstRow -> example 150.2 & selectLastRow -> example 150 -> selectLastRow > selectFirstRow
            //occurrences address, selectedLevelDot Dot count occurences, countRows -1 as Array, 
            int selectFirstRow = -1, selectLastRow = -1, selectedLevelDot = -1, countRows = dataGridView1.Rows.Count - 1;

            if (!CheckActiveSheet())
                return;
            if (!GetMinMaxLevel(ref selectFirstRow, ref selectLastRow, ref selectedLevelDot, countRows)) // First, last range seleted/ designate
                return;
            if (!IsCurrentSelect())
            {
                MessageBox.Show("Data not correct selected !", "Select current level");
                return;
            }
            // Excel & Datagrid moving
            MoveUp(selectFirstRow, selectLastRow, selectedLevelDot, countRows);
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (!CheckActiveSheet())
                return;
            SelectRowsBOM(e.RowIndex);

            if (checkBox3.Checked)
            {
                string text = (string)dataGridView1.Rows[e.RowIndex].Cells[0].Value;
                int level = (int)dataGridView1.Rows[e.RowIndex].Cells[2].Value;

                if (level == 0)
                {
                    textBox1.Text = text;
                }
                else if (level >= 1)
                {
                    textBox1.Text = text.Substring(text.LastIndexOf(".") + 1);
                }
                else
                {
                    MessageBox.Show("Nieznany błąd w pobieraniu danych !");
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //address occurrences-  selectFirstRow -> example 150.2 & selectLastRow -> example 150 -> selectLastRow > selectFirstRow
            //occurrences address, selectedLevelDot Dot count occurences, countRows -1 as Array, 
            int selectFirstRow = -1, selectLastRow = -1, selectedLevelDot = -1, countRows = dataGridView1.Rows.Count - 1;

            if (!CheckActiveSheet())
                return;
            if (!GetMinMaxLevel(ref selectFirstRow, ref selectLastRow, ref selectedLevelDot, countRows)) // First, last range seleted/ designate
                return;
            if (!IsCurrentSelect())
            {
                MessageBox.Show("Data not correct selected !", "Select current level");
                return;
            }
            // Excel & Datagrid moving
            MoveDown(selectFirstRow, selectLastRow, selectedLevelDot, countRows);

            /****************************************
            Obliczanie czasu w s
            double start = DateTime.Now.Ticks; 
            xxxxxxxxxxxxxxxxxx
            double end = ((DateTime.Now.Ticks - start) / 10000000); //wynik w sekundach
            MessageBox.Show("time = " + Math.Round(end, 5) + " s"); //Wyświetla wynik w okienku
            ****************************************/
        }

        private void comboBox1_Click(object sender, EventArgs e)
        {
            if (!CheckActiveSheet())
                return;

            if (currentWorkSheet == null)   // gdy ktos załaduje niezgodny arkusz inny niż IPL i by chciał go używać
            {
                MessageBox.Show("Select current Sheet !!", "Select");
                return;
            }
            else if (xlApp.ActiveSheet.Name != currentWorkSheet.Name)
            {
                MessageBox.Show("Select current Sheet !!", "Select");
                return;
            }

            comboBox1.Items.Clear();
            foreach (Excel.Workbook wb in xlApp.Workbooks) // sprawdzam aktywne wb
            {
                comboBox1.Items.Add(wb.Name);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int selIndex = comboBox1.SelectedIndex, bomFirstRow = 0, bomLastRow = 0, lastRowIPL = 0;
            Excel.Worksheet bomSheet = null;
            Excel.Range bomRange = null, pasteRange=null;

            if (selIndex >= 0)
            {
                bomSheet = (Excel.Worksheet)xlApp.Workbooks.get_Item(selIndex + 1).Sheets.get_Item(1);
            }
            else
            {
                MessageBox.Show("Problem z ustawieniem właściwego sheeta");
                return;
            }

            try
            {
                bomLastRow = bomSheet.Range["B" + bomSheet.Rows.Count].End[Excel.XlDirection.xlUp].Row; // Return last cell in column
                bomRange = bomSheet.get_Range("B1", "B" + bomLastRow.ToString());
                // W sumie zbędne, jakby ktoś wybrał zly arkusz
                bomRange = bomRange.Find("Description", Missing.Value, Excel.XlFindLookIn.xlValues,
                        Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows,
                        Excel.XlSearchDirection.xlNext, false,
                        Missing.Value, Missing.Value);

                if (bomRange == null) //profilaktycznie jakby nie było nagłowka Description 
                {
                    bomSheet = null;
                    MessageBox.Show("Title text 'Description' not found in BOM cells.\nData not correct\nLook on Excel BOM- column $B.\nExcel data check again.", "Sheet check");
                    return;
                }
                bomFirstRow = bomRange.Row; //zwraca adres znalezionego tekstu

                //******
                lastRowIPL = xlWorkSheet.Range["C" + xlWorkSheet.Rows.Count].End[Excel.XlDirection.xlUp].Row; // Return last cell in column

                DialogResult dialogResult = MessageBox.Show("YES (warning overwrite- delete data before use) adding data to: " + xlWorkSheet.Name
                    + " after column $C, row contain title 'DESCRIPTION'\nNO adding data to last row" + xlWorkSheet.Name + "\tY/n?: \t ",
                    "Be careful when you select sheet !", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (dialogResult == DialogResult.Yes)
                {
                    pasteRange = xlWorkSheet.get_Range("C1", "C" + lastRowIPL.ToString());
                    pasteRange = pasteRange.Find("DESCRIPTION", Missing.Value, Excel.XlFindLookIn.xlValues,
                            Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows,
                            Excel.XlSearchDirection.xlNext, false,
                            Missing.Value, Missing.Value);

                    if (pasteRange == null) //if null abort
                    {
                        MessageBox.Show("Title text 'DESCRIPTION' not found in cells.\nLook on Excel- column $C.\nPlease Excel data check again.", "Active sheet check");
                        return;
                    }
                    lastRowIPL = pasteRange.Row;
                }
                else if (dialogResult == DialogResult.Cancel)
                {
                    //MessageBox.Show("PAPAPA");
                    return;
                }

                // kopiowanie
                lastRowIPL = lastRowIPL + 1; //musi być + 1 
                bomFirstRow = bomFirstRow + 1;
                //
                bomRange = bomSheet.get_Range("A" + bomFirstRow.ToString() + ":G" + bomLastRow.ToString());
                pasteRange = xlWorkSheet.get_Range("B" + lastRowIPL.ToString() + ":H" + lastRowIPL.ToString());
                bomRange.Copy(pasteRange);
                bomRange = bomSheet.get_Range("H" + bomFirstRow.ToString() + ":H" + bomLastRow.ToString());
                pasteRange = xlWorkSheet.get_Range("J" + lastRowIPL.ToString() + ":J" + lastRowIPL.ToString());
                bomRange.Copy(pasteRange);

                lastRowIPL = lastRowIPL - bomFirstRow + bomLastRow; // obliczanie do wstawienia X 
                xlWorkSheet.get_Range("B" + lastRowIPL.ToString() + ":B" + lastRowIPL.ToString()).Value = "X";
                // koniec kopiowanie

                if (!GetSheet())
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel Check data.\n " + ex.ToString(), this.Name, MessageBoxButtons.OKCancel, MessageBoxIcon.Stop);
            }
            finally
            {
                bomSheet = null;
                bomRange = null;
                pasteRange = null;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int firstRow = 0, lastRow = 0;
            Excel.Range range = null;

            if (!CheckActiveSheet())
                return;

            if (currentWorkSheet == null)   // gdy ktos załaduje niezgodny arkusz inny niż IPL i by chciał go używać
            {
                MessageBox.Show("Select current Sheet !!", "Select");
                return;
            }
            else if (xlApp.ActiveSheet.Name != currentWorkSheet.Name)
            {
                MessageBox.Show("Select current Sheet !!", "Select");
                return;
            }

            try
            {
                lastRow = xlWorkSheet.Range["C" + xlWorkSheet.Rows.Count].End[Excel.XlDirection.xlUp].Row; // Return last cell in column
                range = xlWorkSheet.get_Range("C1", "C" + lastRow.ToString());
                // description może być małymi 
                range = range.Find("DESCRIPTION", Missing.Value, Excel.XlFindLookIn.xlValues,
                        Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows,
                        Excel.XlSearchDirection.xlNext, false,
                        Missing.Value, Missing.Value);

                if (range == null) //profilaktycznie jakby nie było nagłowka Description 
                {
                    MessageBox.Show("Title text 'Description' not found in BOM cells.\nData not correct\nLook on Excel BOM- column $C.\nExcel data check again.", "Sheet check");
                    return;
                }
                firstRow = range.Row + 1; //zwraca adres znalezionego tekstu
                range = xlWorkSheet.get_Range(firstRow.ToString() + ":" + lastRow.ToString(), Type.Missing);

                DialogResult dialogResult = MessageBox.Show("Excel data delete from " + xlWorkSheet.Name
                    + "\nRange: " + range.Address + "\tY/n?",
                    "Are you sure ?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dialogResult == DialogResult.Yes)
                {
                    if (string.IsNullOrEmpty(xlWorkSheet.get_Range("C" + firstRow.ToString() + ":C" + firstRow.ToString()).Value))
                    {
                        MessageBox.Show("Row is empty !", "Empty row !");
                        return;
                    }
                    else
                        range.Delete(Missing.Value);
                    GetSheet();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel Check data.\n " + ex.ToString(), this.Name, MessageBoxButtons.OKCancel, MessageBoxIcon.Stop);
            }
            finally
            {
                range = null;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (currentWorkSheet != null)
            {
                currentWorkSheet.Activate(); //dziala ;)
                //FillDataGrid(); 
                //xlTemSheet = currentWorkSheet;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Excel.Workbook existWorkBook = null;
            string shortFile, shortTemFile;
            bool isFileOpen = false;

            if (!CheckActiveSheet())
                return;

            if (currentWorkSheet == null)   // gdy ktos załaduje niezgodny arkusz inny niż IPL i by chciał go używać
            {
                MessageBox.Show("Select current Sheet !!", "Select");
                return;
            }
            else if (xlApp.ActiveSheet.Name != currentWorkSheet.Name)
            {
                MessageBox.Show("Select current Sheet !!", "Select");
                return;
            }

            // ^^^^ Load path to the template ^^^^
            if (checkBox1.Checked == true || temPath == "" || !File.Exists(temPath))
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    temPath = openFileDialog1.FileName.ToString();
                    Settings.Default.settingTempPath = temPath;
                    Settings.Default.Save();
                }
                else
                {
                    MessageBox.Show("Select Excel template !\n", this.Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            // ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

            shortTemFile = Path.GetFileNameWithoutExtension(temPath);

            // ***** sprawdzam aktywne workbooki ****
            foreach (Excel.Workbook wb in xlApp.Workbooks) // sprawdzam aktywne wb
            {
                shortFile = (Path.GetFileNameWithoutExtension(wb.FullName));
                if (shortTemFile == shortFile)
                {
                    isFileOpen = true;
                    if ((xlTemWorkBook == null || xlTemSheet == null) && isFileOpen)
                        existWorkBook = wb;
                    break;
                }
            }
            // ********************************

            // ##########################
            if ((xlTemWorkBook == null || xlTemSheet == null) && isFileOpen) // zamykam plik gdy template pusty i jest otwarty skoroszyt
            {
                if (MessageBox.Show("File " + shortTemFile
                    + " is open !\nYes- The template will closed (not will saved, pop-up disable) !!\n"
                    + "No- The template normally closed !", "Close the "
                    + shortTemFile
                    + " or select other.",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Stop) == DialogResult.Yes)
                {
                    xlTemWorkBook = existWorkBook;
                    xlTemSheet = (Excel.Worksheet)xlTemWorkBook.Worksheets.get_Item(1);
                    xlApp.WorkbookBeforeClose += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(xlApp_WorkbookBeforeClose);
                    CloseTemplate();
                }
                return;
            }
            else if ((xlTemWorkBook != null && xlTemSheet != null) && checkBox1.Checked == true && !isFileOpen) // change template, !isFileOpen gdy zamknięty
            {
                // MessageBox.Show("Zmienia template, załadowane ale template zamkniety "); 
                CloseTemplate(); // zamyka template
            }
            else if ((xlTemWorkBook != null && xlTemSheet != null) && checkBox1.Checked == true && isFileOpen) // trzeba zatrzymać i 
            {
                // MessageBox.Show("W nastepnym kroku Ładuje plik o tej samej nazwie Reopen");
                MessageBox.Show("The template: " + shortTemFile
                    + " is reopening !", shortTemFile,
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                CloseTemplate();
            }
            // ##########################

            // ^^^^ Run template ^^^^
            if (checkBox1.Checked == true && !isFileOpen || xlTemWorkBook == null || xlTemSheet == null)
            {
                try
                {
                    xlTemWorkBook = xlApp.Workbooks.Open(temPath,
                        0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                        true, false, 0, true, false, false);

                    xlTemWorkBook.Windows[1].Visible = false;
                    xlTemSheet = (Excel.Worksheet)xlTemWorkBook.Worksheets.get_Item(1);
                    xlApp.WorkbookBeforeClose += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(xlApp_WorkbookBeforeClose);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Template problem, run Excel again, check processes" + ex.ToString());
                }
            }
            // ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

            // ********************************######
            if (checkBox2.Checked == true)
            {
                xlTemWorkBook.Windows[1].Visible = true;
                if (MessageBox.Show("Are you sure to close the template: " + shortTemFile + " ?\n"
                    + "Yes- The template will closed (not will saved, pop-up disable) !!\n"
                    + "No-  Save & close a document as template !", shortTemFile,
                    MessageBoxButtons.YesNo, MessageBoxIcon.Stop) == DialogResult.Yes)
                {
                    CloseTemplate();
                }
                return;
            }
            // uzyc po testach
            // else if (xlTemWorkBook.Windows[1].Visible == true) // gdy jest widoczne okno to przełacza na niewidoczne
            else
                xlTemWorkBook.Windows[1].Visible = false; // dodać jak co //
            // ********************************######

            // ##########################
            try
            {
                if (xlTemWorkBook != null && xlTemSheet != null && (lastRowSheet - firstRowSheet >= 0))
                {
                    int dgLevel = 0, level = 0;
                    // ----------------------------------------------------
                    if (lastRowSheet - firstRowSheet >= 1)
                    {
                        {
                            for (int i = firstRowSheet; i <= lastRowSheet - 1; i++)
                            {
                                dgLevel = i - firstRowSheet;
                                level = (int)dataGridView1.Rows[dgLevel].Cells[2].Value;

                                if (IsBOM(dgLevel))
                                {
                                    xlTemSheet.get_Range((2 * level + 3) + ":" + (2 * level + 3)).Copy(Missing.Value);
                                    xlWorkSheet.get_Range(i + ":" + i).PasteSpecial(Excel.XlPasteType.xlPasteFormats,
                                        Excel.XlPasteSpecialOperation.xlPasteSpecialOperationAdd, false, false);
                                }
                                else
                                {
                                    xlTemSheet.get_Range((2 * level + 4) + ":" + (2 * level + 4)).Copy(Missing.Value);
                                    xlWorkSheet.get_Range(i + ":" + i).PasteSpecial(Excel.XlPasteType.xlPasteFormats,
                                        Excel.XlPasteSpecialOperation.xlPasteSpecialOperationAdd, false, false);
                                }
                            }
                        }
                    }
                    xlTemSheet.get_Range("1:1").Copy(Missing.Value);
                    xlWorkSheet.get_Range(lastRowSheet + ":" + lastRowSheet).PasteSpecial(Excel.XlPasteType.xlPasteFormats,
                        Excel.XlPasteSpecialOperation.xlPasteSpecialOperationAdd, false, false);
                }
                else
                {
                    MessageBox.Show("Problem with reconciliationing !");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem with paint format !\n" + ex.ToString());
            }
            // ##########################
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (Settings.Default.settingTempPath != "")
                temPath = Settings.Default.settingTempPath;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            CloseTemplate();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            int selectFirstRow = -1, selectLastRow = -1, selectedLevelDot = -1, countRows = dataGridView1.Rows.Count - 1,
                startPos = -1, endPos = -1, levelPos = -1;
            string posValue = "";

            if (!CheckActiveSheet())
                return;

            if (currentWorkSheet == null)   // gdy ktos załaduje niezgodny arkusz inny niż IPL i by chciał go używać
            {
                MessageBox.Show("Select current Sheet !!", "Select");
                return;
            }
            else if (xlApp.ActiveSheet.Name != currentWorkSheet.Name)
            {
                MessageBox.Show("Select current Sheet !!", "Select");
                return;
            }

            if (!GetMinMaxLevel(ref selectFirstRow, ref selectLastRow, ref selectedLevelDot, countRows)) // First, last range seleted/ designate
                return;
            if (!IsCurrentSelect())
            {
                MessageBox.Show("Data not correct selected !", "Select current level");
                return;
            }

            // **************
            levelPos = (int)dataGridView1.Rows[selectLastRow].Cells[2].Value;
            for (int i = selectLastRow; i >= selectFirstRow; i--)
            {
                //endPos = posValue.Length;
                posValue = dataGridView1.Rows[i].Cells[0].Value.ToString();
                startPos = IndexOfDots(posValue, levelPos) + 1; // jak wychodzi -1 to start od 0 bo +1
                endPos = IndexOfDots(posValue, levelPos + 1);
                //if (startPos == -1)
                //  startPos = 0;
                if (endPos == -1)
                    endPos = posValue.Length;
                posValue = posValue.Remove(startPos, endPos - startPos).Insert(startPos, textBox1.Text);
                dataGridView1.Rows[i].Cells[0].Value = posValue;
                xlWorkSheet.get_Range("B" + (firstRowSheet + i).ToString() + ":B" + (firstRowSheet + i).ToString()).Value = "'" + posValue;
            }
            // **************
        }

        private void checkBox3_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                button7.Enabled = true;
                textBox1.Enabled = true;
                //MessageBox.Show("True");
            }
            else
            {
                button7.Enabled = false;
                textBox1.Enabled = false;
                textBox1.Text = "Select row;)";
                //MessageBox.Show("False");
            }
        }

        private void checkBox4_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
                this.TopMost = true;
            else
                this.TopMost = false;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            AboutBox1 a = new AboutBox1();
            a.Show();
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            MessageBox.Show("Keypress on datagrid not implemented yet !", "No implemented !", MessageBoxButtons.OK);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (xlTemWorkBook == null || xlTemSheet == null)
                MessageBox.Show("The template isn't set !!", "The template isn't set !!");
            CloseTemplate();
        }
    }
}
