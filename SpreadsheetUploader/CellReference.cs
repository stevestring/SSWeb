using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetUploader
{
    public class CellReference
    {
        public string FullReference { get; set; }

        public CellReference(string fullReference)
        {
            FullReference = fullReference;
        }

        public string FullColumn
        {
            get
            {
                string retVal = "";

                if (FullReference.Substring(0, 1) == "$")
                {
                    retVal = FullReference.Substring(0, 2);
                }
                else
                {
                    retVal = retVal = FullReference.Substring(0, 1);
                }
                return retVal;
            }
        }

        public string FullRow
        {
            get
            {
                string retVal = "";

                if (FullReference.Substring(0, 1) == "$")
                {
                    retVal = FullReference.Substring(2, FullReference.Length - 2);
                }
                else
                {
                    retVal = FullReference.Substring(1, FullReference.Length - 1);
                }
                return retVal;
            }
        }

        public string Row
        {
            get
            {
                return FullRow.Replace("$", "");
            }
        }
        public string Column
        {
            get
            {
                return FullColumn.Replace("$", "");
            }
        }

        public bool IsRelativeRow
        {

            get
            {
                return Row != FullRow;
            }
        }
        public bool IsRelativeColumn
        {

            get
            {
                return Column != FullColumn;
            }
        }

        public int ColumnIndex
        {
            get
            {
                return char.Parse(Column.ToUpper()) - 64;
            }
        }
        public int RowIndex
        {
            get
            {
                return int.Parse(Row);
            }
        }
    }
}
