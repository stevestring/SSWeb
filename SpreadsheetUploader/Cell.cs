using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetUploader
{
    public class Cell
    {
        public int row;
        public int column;
        public string cellAddress;
        public string value;
        public string formula;
        public int dataType;
        public string format;
        public bool readOnly;
        public bool isDerived;
        public bool empty;
        public bool hasDependent;
        public int width;
        public int height;
        public string FontFamily;
        public double FontSize;
        public string ForeColor;
        public string BackColor;
        public bool Underline;
        public bool Bold;
        public bool Italics;
        public bool isSharedForumla;
        public int? sharedForumlaIndex;
        public string fullFormula;
    

    }
}
