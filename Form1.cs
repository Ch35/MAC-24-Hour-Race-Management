using System;
using System.Threading;
using System.Drawing;
using System.IO;
using System.Windows;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Linq;

namespace MAC_24_Hour_Race_Management
{
    public partial class Form1 : Form
    {
        /// <summary>
        /// Start and end date formats
        /// </summary>
        const string DATEFORMAT = "yy/MM/dd HH:mm:ss";

        /// <summary>
        /// Number of decimal points rounded to
        /// </summary>
        const int DECIMALPOINTS = 3;

        /// <summary>
        /// Whether the race has started
        /// </summary>
        private bool started = false;

        /// <summary>
        /// The current path for the uploaded Excel document
        /// </summary>
        private string path;

        /// <summary>
        /// The current Excel document we are working with
        /// </summary>
        private IWorkbook book;

        /// <summary>
        /// Lap counter - column index
        /// </summary>
        private int lapIndex;

        /// <summary>
        /// Handicap lap counter - column index
        /// </summary>
        private int hlapIndex;

        /// <summary>
        /// Start time - column index
        /// </summary>
        private int startIndex;

        /// <summary>
        /// Last lap end time - column index
        /// </summary>
        private int endIndex;

        /// <summary>
        /// Handicap multiplier - column index
        /// </summary>
        private int handicapIndex;

        /// <summary>
        /// A list containing the end time for each row
        /// </summary>
        private List<String> lastEndTime = new List<String>();

        public Form1()
        {
            InitializeComponent();

            LabelTime.Text = DateTime.Now.ToString("HH:mm:ss");
        }

        /// <summary>
        /// Imports the provided Excel Workbook
        /// Creates lap buttons in the first column and an undo button on the final column
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnImport_Click(object sender, EventArgs e)
        {
            // Open file dialog box, allowing the user to path to the desired Excel file
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                //openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Excel XSSF (*.xlsx)|*.xlsx|Excel HSSF (*.xls)|*.xls";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Get the path of specified file - load Excel workbook
                    path = openFileDialog.FileName;
                    book = ReadWorkbook(path);

                    // Load the first sheet from this workbook
                    ISheet sheet = book.GetSheetAt(0);

                    // No sheet found!
                    if (sheet == null)
                    {
                        System.Windows.MessageBox.Show("Excel file does not contain any sheets.", "Excel Read Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }

                    // Fetch no. of col/rows in sheet
                    int rowCount = sheet.LastRowNum;
                    int colCount = sheet.GetRow(0).LastCellNum;

                    // Clear existing rows/columns
                    dataGridView.Rows.Clear();
                    dataGridView.Columns.Clear();

                    // Lap button header
                    dataGridView.Columns.Add("", "");

                    // Reset indices
                    handicapIndex = startIndex = endIndex = lapIndex = hlapIndex = 0;

                    // Set column values
                    for (int c = 0; c < colCount; c++)
                    {
                        ICell cell = sheet.GetRow(0).GetCell(c);
                        string cellVal = cell.ToString();
                        int cellIndex = c + 1;

                        // Store all the required indices
                        switch (cellVal)
                        {
                            case "Handicap":
                                handicapIndex = cellIndex;
                                break;
                            case "Start":
                                startIndex = cellIndex;
                                break;
                            case "Last Lap End":
                                endIndex = cellIndex;
                                break;
                            case "Laps":
                                lapIndex = cellIndex;
                                break;
                            case "Handicap Laps":
                                hlapIndex = cellIndex;
                                break;
                        }

                        dataGridView.Columns.Add(c.ToString() + cellVal, cellVal);
                    }

                    // add an undoLap column
                    dataGridView.Columns.Add("", "");

                    // Missing a required column - They should add up to or be more than 10
                    if (handicapIndex == 0 || startIndex == 0 || endIndex == 0 || lapIndex == 0 || hlapIndex == 0)
                    { 
                        // reset everything
                        dataGridView.Rows.Clear();
                        dataGridView.Columns.Clear();
                        book = null;
                        path = null;

                        string error = "Missing required column headers ['Handicap', 'Start', 'Last Lap End', 'Laps', 'Handicap Laps']";
                        System.Windows.MessageBox.Show(error, "Excel Read Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    // we can export some data now
                    if (book != null)
                    {
                        BtnExport.Visible = true;
                        BtnStart.Visible = true;
                        //BtnSaldanha.Visible = true;
                        started = false;
                    }

                    // Set row values
                    for (int r = 1; r <= rowCount; r++)
                    {
                        IRow row = sheet.GetRow(r);

                        // Row cant have empty values for all cells
                        bool empty = true;
                        foreach (var cell in row.Cells)
                        {
                            // cell exists and isnt an empty string
                            if (cell != null && cell.ToString() != "")
                            {
                                empty = false;
                                break;
                            }
                        }

                        // Row is empty, no need to add it
                        if (empty == true)
                        {
                            continue;
                        }

                        int rowIndex = dataGridView.Rows.Add();
                        int colIndex = dataGridView.Columns.Count - 1;

                        // Create buttoncel for the LAP button
                        //dataGridView.Rows[rowIndex].Cells[0] = createButtonCell("Lap", Color.FromArgb(0, 155, 0));
                        dataGridView.Rows[rowIndex].Cells[0].Value = "Lap";
                        dataGridView.Rows[rowIndex].Cells[0].Style.BackColor = Color.FromArgb(2, 1, 51);

                        // Create a buttoncell for the undoLap 
                        //var btn = createButtonCell("Undo", Color.FromArgb(2, 1, 51));
                        dataGridView.Rows[rowIndex].Cells[colIndex].Value = "Undo";
                        dataGridView.Rows[rowIndex].Cells[colIndex].Style.BackColor = Color.FromArgb(2, 1, 51);

                        // set column values
                        for (int c = 0; c < colIndex - 1; c++)
                        {
                            ICell cell = row.GetCell(c);

                            // Need to default this cell to 0 if Excell cell is empty
                            int cellIndex = c + 1;
                            string value = "";

                            if (cell != null)
                            {
                                value = cell.ToString();
                            }

                            // empty Lap cell
                            if (value == "" && (cellIndex == lapIndex || cellIndex == hlapIndex))
                            {
                                value = "0";
                            }

                            dataGridView.Rows[rowIndex].Cells[cellIndex].Value = value;
                        }
                    }

                }
            }

            dataGridView.CurrentCell = null;
        }

        /// <summary>
        /// Reads Excel workbook from string path
        /// </summary>
        /// <param name="path"></param>
        /// <returns>IWorkbook</returns>
        private IWorkbook ReadWorkbook(string path)
        {
            IWorkbook book;

            try
            {
                FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                // Try to read workbook as XLSX:
                try
                {
                    book = new XSSFWorkbook(fs);
                }
                catch
                {
                    book = null;
                }

                // If reading fails, try to read workbook as XLS:
                if (book == null)
                {
                    book = new HSSFWorkbook(fs);
                }

                // Clear the FileStream
                fs.Close();
                fs.Dispose();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message, "Excel Read Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }

            return book;
        }

        /// <summary>
        /// Starts the race. The race cannot start again once it has started.
        /// Writes the current time to the "Race Start" column
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnStart_Click(object sender, EventArgs e)
        {
            started = true;

            DataGridViewRow lastRow = dataGridView.Rows[dataGridView.Rows.Count - 1];

            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                row.Cells[startIndex].Value = DateTime.Now.ToString(DATEFORMAT);
            }
        }

        /// <summary>
        /// Updates the label indicator for the current time
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Timer_Tick(object sender, EventArgs e)
        {
            LabelTime.Text = DateTime.Now.ToString("HH:mm:ss");
        }

        /// <summary>
        /// Exports the DataGridView to an Excel document
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnExport_Click(object sender, EventArgs e)
        {
            SaveFileDialog fileDialog = new SaveFileDialog();
            fileDialog.Filter = "Excel XSSF|*.xlsx";

            if (fileDialog.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }

            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();
            IRow rowHead = sheet.CreateRow(0);

            // Set headers (Ignoring button columns)
            for(int i = 1; i < dataGridView.Columns.Count - 1; i++)
            {
                var col = dataGridView.Columns[i];

                if(col.HeaderText != null)
                {
                    rowHead.CreateCell(i-1, CellType.String).SetCellValue(col.HeaderText);
                }
            }

            // Set row data
            for (int i = 0; i < dataGridView.Rows.Count; i++)
            {
                var row = dataGridView.Rows[i];
                IRow sheetRow = sheet.CreateRow(i + 1);

                // iterate through row cells (Excluding button columns)
                for(int j = 1; j < row.Cells.Count - 1; j++)
                {
                    var cell = row.Cells[j];

                    if (cell.Value != null)
                    {
                        sheetRow.CreateCell(j-1, CellType.String).SetCellValue(cell.Value.ToString());
                    }
                }
            }

            // autosize all columns
            int c = 0;
            foreach (DataGridViewColumn col in dataGridView.Columns)
            {
                sheet.AutoSizeColumn(c);
                c++;
            }

            // Finally save the file
            using (FileStream stream = File.OpenWrite(fileDialog.FileName))
            {
                workbook.Write(stream);
                stream.Close();
            }

            // Memory garbage collector
            GC.Collect();
        }

        /// <summary>
        /// Creates a lap
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // Ignore header row
            if (e.RowIndex == -1)
            {
                return;
            }

            DataGridViewRow row = dataGridView.Rows[e.RowIndex];

            // Not a button cell
            if (e.ColumnIndex != 0 && e.ColumnIndex != (row.Cells.Count - 1))
            {
                return;
            }

            if (started == false)
            {
                string msg = "Please start the race for this class before incrementing/undoing laps.";
                System.Windows.MessageBox.Show(msg, "Race not started", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // Lap Cell
            if(e.ColumnIndex == 0)
            {
                lap(e);
            }
            // Undo Lap Cell
            else
            {
                undolap(e);
            }
        }

        /// <summary>
        /// Incrememnts a lap for the given row
        /// Writes the lap time to the "last lap end"
        /// </summary>
        /// <param name="e"></param>
        private void lap(DataGridViewCellEventArgs e)
        {
            // Fetch the current row
            DataGridViewRow row = dataGridView.Rows[e.RowIndex];

            DataGridViewCell lapCell = row.Cells[lapIndex];
            DataGridViewCell hlapCell = row.Cells[hlapIndex];
            DataGridViewCell endCell = row.Cells[endIndex];

            // Calculate lap cell
            int lapCount;

            lapCount = (lapCell.Value == null) ? 0 : Int32.Parse(lapCell.Value.ToString());
            lapCount++;

            // preserve the previous cell value
            //lastEndTime[e.RowIndex] = endCell.Value.ToString();
            lastEndTime.Insert(e.RowIndex, endCell.Value.ToString());

            // set the current time to the last lap
            string endTime = DateTime.Now.ToString(DATEFORMAT);
            endCell.Value = endTime;

            // Increment laps (including handicap laps)
            lapCell.Value = lapCount.ToString();
            hlapCell.Value = calculateHandicapLap(row, lapCount).ToString();
        }

        /// <summary>
        /// Undos a lap
        /// </summary>
        /// <param name="e"></param>
        private void undolap(DataGridViewCellEventArgs e)
        {
            // Fetch the current row
            DataGridViewRow row = dataGridView.Rows[e.RowIndex];

            // Previous value exists, set the lap end time to this
            if (lastEndTime.ElementAtOrDefault(e.RowIndex) == null)
            {
                System.Windows.MessageBox.Show("You are only able to undo once for each lap", "Cannot Undo", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            DataGridViewCell endCell = row.Cells[endIndex];
            DataGridViewCell lapCell = row.Cells[lapIndex];
            DataGridViewCell hlapCell = row.Cells[hlapIndex];

            //// Previous value exists, set the lap end time to this
            //if (lastEndTime.ElementAtOrDefault(e.RowIndex) != null)
            //{
            endCell.Value = lastEndTime[e.RowIndex];
            lastEndTime[e.RowIndex] = null;
            //}

            // deincrement 
            int lapCount = Int32.Parse(lapCell.Value.ToString());
            if(lapCount > 0)
            {
                lapCount--;
            }
            lapCell.Value = lapCount.ToString();

            // update handicaplaps
            hlapCell.Value = calculateHandicapLap(row, lapCount).ToString();
        }

        /// <summary>
        /// Calculates the number of handicap laps this boat should have based on the lapcount and row
        /// </summary>
        /// <param name="row"></param>
        /// <param name="lapCount"></param>
        /// <returns></returns>
        private double calculateHandicapLap(DataGridViewRow row, int lapCount)
        {
            DataGridViewCell handicapCell = row.Cells[handicapIndex];
            string handicapValue = handicapCell.Value.ToString();
            char separator = Convert.ToChar(Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator);

            // cell does not use the system number decimal separator
            if (handicapValue.Contains(separator.ToString()) == false)
            {
                char used = (separator == ',') ? '.' : ','; // used is the inverse of the separator
                handicapValue = handicapValue.Replace(used, separator);
            }
            float handicap = float.Parse(handicapValue);

            return Math.Round(lapCount * handicap, DECIMALPOINTS);
        }
    }
}
