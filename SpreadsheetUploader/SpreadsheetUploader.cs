using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using System.Configuration;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Data;

namespace SpreadsheetUploader
{


    public class SpreadsheetUploader
    {

        Dictionary<string, Cell> savedCells;
        Dictionary<string, Cell> allCells;

        private List<string> GetCellReferences(string formula)
        {
            List<string> retVal = new List<string>();

            string regex = @"\$?(?:\bXF[A-D]|X[A-E][A-Z]|[A-W][A-Z]{2}|[A-Z]{2}|[A-Z])\$?(?:104857[0-6]|10485[0-6]\d|1048[0-4]\d{2}|104[0-7]\d{3}|10[0-3]\d{4}|[1-9]\d{1,5}|[1-9])d?\b([:\s]\$?(?:\bXF[A-D]|X[A-E][A-Z]|[A-W][A-Z]{2}|[A-Z]{2}|[A-Z])\$?(?:104857[0-6]|10485[0-6]\d|1048[0-4]\d{2}|104[0-7]\d{3}|10[0-3]\d{4}|[1-9]\d{1,5}|[1-9])d?\b)?";


            Match match = Regex.Match(formula, regex,
                RegexOptions.IgnoreCase);

            while (match.Success)
            {

                //// Finally, we get the Group value and display it.
                string key = match.Groups[0].Value;
                string cellRef = key;// key.Replace("$", "");
                retVal.Add(cellRef);

                formula = formula.Replace(key, "");

                match = Regex.Match(formula, regex,
                RegexOptions.IgnoreCase);
            }

            return retVal;

        }



        private int PutWorksheet(string name)
        {
            SqlConnection conn = new SqlConnection(ConnectionString);
            conn.Open();

            SqlCommand cmd = new SqlCommand("put_worksheet", conn);
            cmd.CommandType = CommandType.StoredProcedure;

            cmd.Parameters.AddWithValue("name", name);

            int retVal = int.Parse(cmd.ExecuteScalar().ToString());

            conn.Close();

            return retVal;
        }

        public int LockWorksheet(int id, Boolean locked)
        {
            SqlConnection conn = new SqlConnection(ConnectionString);
            conn.Open();

            SqlCommand cmd = new SqlCommand("lock_worksheet", conn);
            cmd.CommandType =  CommandType.StoredProcedure;

            cmd.Parameters.AddWithValue("worksheet_id", id);
            cmd.Parameters.AddWithValue("loading_cells", locked);

            int retVal = cmd.ExecuteNonQuery();

            conn.Close();

            return retVal;
        }



        private void PutRow(int worksheetId, Cell cell)
        {
            SqlConnection conn = new SqlConnection(ConnectionString);
            conn.Open();

            SqlCommand cmd = new SqlCommand("put_worksheet_cell", conn);
            cmd.CommandType = CommandType.StoredProcedure;

            cmd.Parameters.AddWithValue("worksheet_id", worksheetId);
            cmd.Parameters.AddWithValue("row", cell.row);
            cmd.Parameters.AddWithValue("column", cell.column);
            cmd.Parameters.AddWithValue("cell_address", cell.cellAddress);
            cmd.Parameters.AddWithValue("value", cell.value);
            cmd.Parameters.AddWithValue("formula", cell.formula);

            cmd.Parameters.AddWithValue("data_type", cell.dataType);
            cmd.Parameters.AddWithValue("format", cell.format);
            cmd.Parameters.AddWithValue("read_only", cell.readOnly);
            cmd.Parameters.AddWithValue("is_derived", cell.isDerived);
            cmd.Parameters.AddWithValue("has_dependent", cell.hasDependent);

            cmd.ExecuteNonQuery();

            conn.Close();

        }

        private void DeleteSheet(int worksheetId)
        {
            SqlConnection conn = new SqlConnection(ConnectionString);

            conn.Open();

            SqlCommand cmd = new SqlCommand("delete_worksheet", conn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("worksheet_id", worksheetId);
            cmd.ExecuteNonQuery();
            conn.Close();

        }



        private Worksheet GetNextWorksheet()
        {
            Worksheet retVal = new Worksheet();
            SqlConnection conn = new SqlConnection(ConnectionString);

            conn.Open();


            SqlCommand cmd = new SqlCommand("get_worksheet_stub", conn);
            cmd.CommandType = CommandType.StoredProcedure;
            SqlDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                retVal.WorksheetID = int.Parse(dr["worksheet_id"].ToString());
                retVal.FileName = dr["file_name"].ToString();
            }

            conn.Close();

            return retVal;
        }

        public void SetWorkSheetComplete(int worksheetId)
        {
            Worksheet retVal = new Worksheet();
            SqlConnection conn = new SqlConnection(ConnectionString);

            conn.Open();


            SqlCommand cmd = new SqlCommand("set_worksheet_complete", conn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("worksheet_id", worksheetId);
            cmd.ExecuteNonQuery();
            conn.Close();

        }


        private int GetWorksheetByName(string sheetName)
        {
            int retVal = -1;
            SqlConnection conn = new SqlConnection(ConnectionString);

            conn.Open();


            SqlCommand cmd = new SqlCommand("get_worksheet_by_name", conn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("name", sheetName);
            SqlDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                retVal = int.Parse(dr["worksheet_id"].ToString());
            }

            conn.Close();

            return retVal;
        }

 
        public void LoadSpreadsheet(int worksheetId, string fileName)
        {

            // Create new stopwatch
            System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();

            // Begin timing
            stopwatch.Start();


            string cellAddress = "";
            string sheetName = Path.GetFileName(fileName).Replace(".xlsx", "");
            sheetName = Path.GetFileName(fileName).Replace(".xls", "");

            savedCells = new Dictionary<string, Cell>();
            allCells = new Dictionary<string, Cell>();

            string workbookPath = fileName;

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {

                var sheets = spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>();


                int? colIndex = 0;
                foreach (Sheet sheet in sheets)
                {
                    WorksheetPart worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(sheet.Id);
                    DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet = worksheetPart.Worksheet;

                    SharedStringTablePart sharedStringTablePart1;
                    if (spreadsheetDocument.WorkbookPart.GetPartsCountOfType
                        <SharedStringTablePart>() > 0)
                    {
                        sharedStringTablePart1 =
                            spreadsheetDocument.WorkbookPart.GetPartsOfType
                            <SharedStringTablePart>().First();
                    }
                    else//This is only needed to compile.  If we don't have string table that means there are no strings
                    {                          
                        sharedStringTablePart1 =
                                spreadsheetDocument.WorkbookPart.AddNewPart
                                <SharedStringTablePart>();
                    }



                    var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>();
                    foreach (var r in rows)
                    {
                        if (r.RowIndex <= 30)//30 rows
                        {
                            foreach (DocumentFormat.OpenXml.Spreadsheet.Cell rCell in r.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>())
                            {
                                colIndex = GetColumnIndex(rCell.CellReference.ToString());
                                if (colIndex <= 12)
                                {
                                    Cell c = new Cell();                                    
                                    
                                    cellAddress = rCell.CellReference.ToString().Replace("$", "");

                                    c.readOnly = true;//until updated
                                    c.isDerived = false;
                                    c.row = int.Parse(r.RowIndex.ToString());
                                    c.column = (int)colIndex;
                                    c.cellAddress = cellAddress;
                                    c.value = rCell.CellValue.Text;

                                    if (cellAddress =="A3")
                                    {
                                        Console.WriteLine("A3");
                                    }

                                    //Add Formulas
                                    if (rCell.CellFormula != null)
                                    {
                                        if (rCell.CellFormula.SharedIndex != null)
                                        {
                                            c.sharedForumlaIndex =
                                                    int.Parse(rCell.CellFormula.SharedIndex.ToString());
                                            if (rCell.CellFormula.Text != "")
                                            {
                                                c.isSharedForumla = true;
                                            }
                                        }

                                        c.fullFormula = rCell.CellFormula.Text;
                                        c.formula = "=" + //Add =
                                            (string)rCell.CellFormula.Text.Replace("$", "");//Remove any $


                                        string regex = @"\$?(?:\bXF[A-D]|X[A-E][A-Z]|[A-W][A-Z]{2}|[A-Z]{2}|[A-Z])\$?(?:104857[0-6]|10485[0-6]\d|1048[0-4]\d{2}|104[0-7]\d{3}|10[0-3]\d{4}|[1-9]\d{1,5}|[1-9])d?\b([:\s]\$?(?:\bXF[A-D]|X[A-E][A-Z]|[A-W][A-Z]{2}|[A-Z]{2}|[A-Z])\$?(?:104857[0-6]|10485[0-6]\d|1048[0-4]\d{2}|104[0-7]\d{3}|10[0-3]\d{4}|[1-9]\d{1,5}|[1-9])d?\b)?";
                                        Match match = Regex.Match(c.formula, regex,
                                            RegexOptions.IgnoreCase);

                                        while (match.Success && c.formula.IndexOf(":") != -1)
                                        {
                                            if (match.Groups[0].Value.IndexOf(":") != -1)
                                            {
                                                //// Finally, we get the Group value and display it.
                                                string key = match.Groups[0].Value;
                                                string cellRef = key.Replace("$", "");

                                                string newCells = RangeToSeries(key);


                                                c.formula = c.formula.Replace(key, newCells);

                                            }
                                            match = Regex.Match(c.formula, regex,
                                            RegexOptions.IgnoreCase);
                                        }

                                    }
                                    else //String (use shared string table)
                                    {

                                        c.value = sharedStringTablePart1.SharedStringTable.ChildElements[int.Parse(c.value)].InnerText;
                                    }

                                    allCells.Add(cellAddress, c);

                                }

                            }

                        }
                    }
                }
                Console.WriteLine();
            }

            ReplaceSharedFormulas(ref allCells);



            Console.WriteLine("Loaded from Excel:" + stopwatch.Elapsed);

            //while (savedCells.Count < 360)
            //{
                foreach (Cell rCell in allCells.Values)
                {
                    if (!savedCells.ContainsKey(rCell.cellAddress))
                    {
                        if (rCell.value == null)//Blank?
                        {
                            savedCells.Add(rCell.cellAddress, rCell);
                        }
                        else if (rCell.formula == null)//Input
                        {
                            rCell.readOnly = true;//until updated (May be label, May be input) 
                            rCell.isDerived = false;
                            savedCells.Add(rCell.cellAddress, rCell);
                        }
                        else //Formula
                        {
                            bool allCellsHere = true;
                            foreach (string s in GetCellReferences((string)rCell.formula))
                            {

                                if (!savedCells.ContainsKey(s))
                                {
                                    allCellsHere = false;
                                    break;
                                }
                                else
                                {
                                    if (savedCells[s].formula == null) //Is an input
                                    {
                                        savedCells[s].empty = false;
                                        savedCells[s].readOnly = false;
                                        savedCells[s].hasDependent = true;
                                    }
                                    //c.empty = false;//this cell impacts others
                                }

                            }

                            if (allCellsHere)
                            {

                                rCell.formula = rCell.formula.Replace("$", "");
                                rCell.readOnly = true;
                                rCell.isDerived = true;
                                savedCells.Add(rCell.cellAddress, rCell);
                            }

                        }
                    }
                }

            //}
            Console.WriteLine("Organized Cells:" + stopwatch.Elapsed);

            Dictionary<string, Cell> savedCellsTrimmed = new Dictionary<string, Cell>();
            List<int> rowsToRemove = new List<int>();

            bool includeRow = false;
            for (int k = 30; k > 0; k--)
            {
                foreach (Cell c in savedCells.Values)
                {
                    if (c.row == k)
                    {
                        if (c.value != "" || c.hasDependent)
                        {
                            includeRow = true;
                            break; //include row;
                        }
                    }

                }
                if (includeRow)
                {
                    break;
                }
                else
                {
                    rowsToRemove.Add(k);
                }
            }

            foreach (Cell c in savedCells.Values)
            {
                if (!rowsToRemove.Contains(c.row))
                {
                    savedCellsTrimmed.Add(c.cellAddress, c);
                }
            }

            savedCells = savedCellsTrimmed;

            DataTable dt = new DataTable();
            dt.Columns.Add("worksheet_id");
            dt.Columns.Add("row");
            dt.Columns.Add("column");
            dt.Columns.Add("cell_address");
            dt.Columns.Add("value");
            dt.Columns.Add("formula");
            dt.Columns.Add("data_type");
            dt.Columns.Add("format");
            dt.Columns.Add("read_only");
            dt.Columns.Add("is_derived");
            dt.Columns.Add("has_dependent");

            dt.Columns.Add("fore_color");
            dt.Columns.Add("back_color");
            dt.Columns.Add("font_family");
            dt.Columns.Add("italics");
            dt.Columns.Add("bold");
            dt.Columns.Add("underline");
            dt.Columns.Add("font_size");



            foreach (Cell c in savedCells.Values)
            {

                DataRow row = dt.NewRow();

                row[0] = worksheetId;
                row[1] = c.row;
                row[2] = c.column;
                row[3] = c.cellAddress;
                row[4] = c.value;
                row[5] = c.formula;
                row[6] = c.dataType;
                row[7] = c.format;
                row[8] = c.readOnly;
                row[9] = c.isDerived;
                row[10] = c.hasDependent;
                row[11] = c.ForeColor;
                row[12] = c.BackColor;
                row[13] = c.FontFamily;
                row[14] = c.Italics;
                row[15] = c.Bold;
                row[16] = c.Underline;
                row[17] = c.FontSize;
                dt.Rows.Add(row);

            }

            SqlConnection cn = new SqlConnection(ConnectionString);
            cn.Open();
            using (SqlBulkCopy copy = new SqlBulkCopy(cn))
            {
                copy.ColumnMappings.Add(0, 0);
                copy.ColumnMappings.Add(1, 1);
                copy.ColumnMappings.Add(2, 2);
                copy.ColumnMappings.Add(3, 3);
                copy.ColumnMappings.Add(4, 4);
                copy.ColumnMappings.Add(5, 5);
                copy.ColumnMappings.Add(6, 6);
                copy.ColumnMappings.Add(7, 7);
                copy.ColumnMappings.Add(8, 8);
                copy.ColumnMappings.Add(9, 9);
                copy.ColumnMappings.Add(10, 10);
                copy.ColumnMappings.Add(11, 11);
                copy.ColumnMappings.Add(12, 12);
                copy.ColumnMappings.Add(13, 13);
                copy.ColumnMappings.Add(14, 14);
                copy.ColumnMappings.Add(15, 15);
                copy.ColumnMappings.Add(16, 16);
                copy.ColumnMappings.Add(17, 17);

                copy.DestinationTableName = "worksheet_cell";
                copy.WriteToServer(dt);
            }

            Console.WriteLine("Loaded DB:" + stopwatch.Elapsed);

        }



        private void ReplaceSharedFormulas(ref Dictionary<string, Cell> allCells)
        {
            foreach (var item in allCells)
            {
                if (item.Value.cellAddress == "J5")
                {
                    Console.Write("J5");
                }



                if (item.Value.sharedForumlaIndex != null
                    && !item.Value.isSharedForumla)//Needs new formula
                {

                    foreach (var item2 in allCells)
                    {
                        if (item2.Value.isSharedForumla &&
                            item2.Value.sharedForumlaIndex == item.Value.sharedForumlaIndex)//Is formulaindex the same?
                        {
                            item.Value.formula = ConvertSharedFormula(item2.Value, item.Value);
                        }
                    }
                }
            }
        }

        private string ConvertSharedFormula(Cell parentCell, Cell childCell)
        {
            string newFormula = parentCell.fullFormula;
            string newReference = "";
            int colOffset = childCell.column - parentCell.column;
            int rowOffset = childCell.row - parentCell.row;

            List<string> origCellReferences = GetCellReferences(newFormula);

            if (childCell.cellAddress == "J5")
            {
                Console.Write("J5");
            }

            foreach (string s in origCellReferences)
            {
                CellReference cr = new CellReference(s);
                int oldcolumn = cr.ColumnIndex; //Char.Parse(s.Substring(0, 1).ToUpper()) - 64;
                int newColumn = oldcolumn;

                if (!cr.IsRelativeColumn)
                {
                    newColumn += colOffset;
                }

                string newColumnString = ((char)(newColumn + 64)).ToString();

                int oldRow = cr.RowIndex;
                string newRowString = cr.Row;

                if (!cr.IsRelativeRow)
                {
                    newRowString = (cr.RowIndex + rowOffset).ToString();
                }
                newReference = newColumnString + newRowString;
                newFormula = newFormula.Replace(s, newReference);
            }


            return "=" + newFormula;
        }



        private static int? GetColumnIndex(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference))
            {
                return null;
            }

            //remove digits
            string columnReference = Regex.Replace(cellReference.ToUpper(), @"[\d]", string.Empty);

            int columnNumber = -1;
            int mulitplier = 1;

            //working from the end of the letters take the ASCII code less 64 (so A = 1, B =2...etc)
            //then multiply that number by our multiplier (which starts at 1)
            //multiply our multiplier by 26 as there are 26 letters
            foreach (char c in columnReference.ToCharArray().Reverse())
            {
                columnNumber += mulitplier * ((int)c - 64);

                mulitplier = mulitplier * 26;
            }

            //the result is zero based so return columnnumber + 1 for a 1 based answer
            //this will match Excel's COLUMN function
            return columnNumber + 1;
        }


        

        public string RangeToSeries(string formula)
        {
            string retVal = "";
            string[] cells = formula.Split(Char.Parse(":"));
            string cell1 = cells[0];
            string cell2 = cells[1];

            string column1 = cells[0].Substring(0, 1);
            string row1 = cells[0].Substring(column1.Length, cell1.Length - column1.Length);
            string column2 = cells[1].Substring(0, 1);
            string row2 = cells[1].Substring(column2.Length, cell2.Length - column2.Length);

            int column1int = LetterToColumnNumber(char.Parse(column1));
            int column2int = LetterToColumnNumber(char.Parse(column2));

            int row1int = int.Parse(row1);
            int row2int = int.Parse(row2);

            int minColumn = 0;
            int maxColumn = 0;

            int minRow = 0;
            int maxRow = 0;

            if (column1int > column2int)
            {
                maxColumn = column1int;
                minColumn = column2int;
            }
            else
            {
                minColumn = column1int;
                maxColumn = column2int;
            }

            if (row1int > row2int)
            {
                maxRow = row1int;
                minRow = row2int;
            }
            else
            {
                minRow = row1int;
                maxRow = row2int;
            }

            List<string> newCells = new List<string>();
            for (int i = minRow; i <= maxRow; i++)
            {

                for (int j = minColumn; j <= maxColumn; j++)
                {
                    newCells.Add(ColumnNumberToLetter(j) + i.ToString());
                }
            }


            foreach (string s in newCells)
            {
                retVal += s + ",";
            }

            retVal = retVal.TrimEnd(Char.Parse(","));

            return retVal;


        }

        // Return the column number for this letter.
        private int LetterToColumnNumber(char letter)
        {
            // See if it's out of bounds.
            if (letter < 'A') return 0;
            if (letter > 'Z') return 25;

            // Calculate the number.
            return (int)letter - (int)'A';
        }

        // Return the letter for this column number.
        private char ColumnNumberToLetter(int number)
        {
            // See if it's out of bounds.
            if (number < 0) return 'A';
            if (number > 25) return 'Z';

            // Calculate the letter.
            return (char)(number + (int)'A');
        }
        public string GetColumnName(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            var value = "";

            if (index >= letters.Length)
                value += letters[index / letters.Length - 1];

            value += letters[index % letters.Length];

            return value;
        }


        private void CopyFile(string fileName)
        {
            File.Copy(@"C:\" + fileName, @"C:\SSLoader\Downloads\" + fileName, true);
        }


 
        public static string ConnectionString
        {
            get
            {
                return ConfigurationManager.ConnectionStrings["ApplicationServices"].ToString();
            }
        }
    }
}
