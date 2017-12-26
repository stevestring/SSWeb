using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Web.Mvc;
using System.Web.Security;
using System.Text.RegularExpressions;
namespace SSWeb.Models
{
    public class Worksheet
    {
        [Required]
        public string Name { get; set; }
        public int WorksheetId { get; set; }
        public string Description { get; set; }
        public string CreatedBy { get; set; }
        public string Owner { get; set; }
        public int Complete { get; set; }
        public bool Private { get; set; }
        public string Format { get; set; }
        public bool Failed { get; set; }
        public string ErrorMessage { get; set; }


        public Dictionary<int, Dictionary<int, Cell>> Cells;

        private string _calcFinal = "";
        public string Calcfinal
        {
            get
            {
                return _calcFinal;
            }
        }

        public void GetWorksheet(int id)
        {

            SqlConnection conn = new SqlConnection(SSWeb.Helpers.Database.ConnectionString);

            conn.Open();

            SqlCommand cmd = new SqlCommand("get_worksheet", conn);
            cmd.Parameters.AddWithValue("worksheet_id", id);
            cmd.CommandType = CommandType.StoredProcedure;

            SqlDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                WorksheetId = int.Parse(dr["worksheet_id"].ToString());
                Name = dr["name"].ToString();
                Description = dr["description"].ToString();
                Complete = int.Parse(dr["complete"].ToString());
                Owner = dr["owner"].ToString();
                Private = bool.Parse(dr["private"].ToString());

                if (dr["error_message"] != DBNull.Value)
                {
                    ErrorMessage = dr["error_message"].ToString();
                    Failed = true;
                }
                else
                {
                    Failed = false;
                }
                //Format = dr["format"].ToString();
                //if (Format == "General") 
                //{
                //    Format = "0.00";
                //}
            }

            conn.Close();            

        }

        public int PutWorksheetStub(string fileName)
        {
            SqlConnection conn = new SqlConnection(SSWeb.Helpers.Database.ConnectionString);
            conn.Open();

            SqlCommand cmd = new SqlCommand("put_worksheet_stub", conn);
            cmd.CommandType = CommandType.StoredProcedure;

            cmd.Parameters.AddWithValue("file_name", fileName);
            cmd.Parameters.AddWithValue("owner", Owner);

            int retVal = int.Parse(cmd.ExecuteScalar().ToString());

            conn.Close();

            return retVal;
        }


        public void UpdateWorksheet(int id, string name, string description, string createdBy, bool isPrivate)
        {
            SqlConnection conn = new SqlConnection(SSWeb.Helpers.Database.ConnectionString);
            conn.Open();

            SqlCommand cmd = new SqlCommand("update_worksheet", conn);
            cmd.CommandType = CommandType.StoredProcedure;

            cmd.Parameters.AddWithValue("worksheet_id", id);
            cmd.Parameters.AddWithValue("name", name);
            cmd.Parameters.AddWithValue("description", description);
            cmd.Parameters.AddWithValue("created_by", createdBy);
            cmd.Parameters.AddWithValue("private", Private);
            //cmd.Parameters.AddWithValue("owner", createdBy);

            cmd.ExecuteNonQuery();

            conn.Close();

        }


        public void DeleteWorksheet(int id)
        {
            SqlConnection conn = new SqlConnection(SSWeb.Helpers.Database.ConnectionString);
            conn.Open();

            SqlCommand cmd = new SqlCommand("delete_worksheet", conn);
            cmd.CommandType = CommandType.StoredProcedure;

            cmd.Parameters.AddWithValue("worksheet_id", id);

            cmd.ExecuteNonQuery();

            conn.Close();

        }
        public void GetWorksheetCells(int id)
        {
            Cells = new Dictionary<int, Dictionary<int, Cell>>(); 
            
            SqlConnection conn = new SqlConnection(SSWeb.Helpers.Database.ConnectionString);
 
            conn.Open();

            SqlCommand cmd = new SqlCommand("get_worksheet_cells", conn);
            cmd.Parameters.AddWithValue("worksheet_id", id);
            cmd.CommandType = CommandType.StoredProcedure;

            SqlDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                Cell c = new Cell();

                c.row = int.Parse(dr["row"].ToString());
                
                c.column = int.Parse(dr["column"].ToString());
                c.cellAddress = dr["cell_address"].ToString();

                if (dr["value"] != DBNull.Value)
                {
                    c.value = dr["value"].ToString();
                }
                if (dr["formula"] != DBNull.Value)
                {
                    c.formula = dr["formula"].ToString();
                }
                //c.dataType = int.Parse(dr["data_type"].ToString());
                //c.format = dr["format"].ToString();

                c.readOnly = Boolean.Parse(dr["read_only"].ToString());
                c.isDerived = Boolean.Parse(dr["is_derived"].ToString());
                c.hasDependent = Boolean.Parse(dr["has_dependent"].ToString());

                if (dr["format"] != null)
                {
                    c.format = dr["format"].ToString();
                }

                //if (dr["font_family"] != null)
                //{
                //    c.FontFamily = dr["font_family"].ToString();
                //}
                //if (dr["font_size"] != null)
                //{
                //    c.FontSize = int.Parse(dr["font_size"].ToString());
                //}
                if (dr["fore_color"] != null)
                {
                    c.ForeColor=dr["fore_color"].ToString();
                }

                if (dr["back_color"] != null)
                {
                    c.BackColor = dr["back_color"].ToString();
                }

                if (dr["underline"] != DBNull.Value)
                {
                    c.Underline = Boolean.Parse(dr["underline"].ToString());
                  
                }
                if (dr["bold"] != DBNull.Value)
                {
                    c.Bold = Boolean.Parse(dr["bold"].ToString());
                }
                if (dr["italics"] != DBNull.Value)
                {
                    c.Italics = Boolean.Parse(dr["italics"].ToString()); 
                }
               
                




                if (!Cells.ContainsKey(c.row))
                {
                    Cells.Add(c.row,new Dictionary<int,Cell>());
                }
                if (!Cells[c.row].ContainsKey(c.column))
                {
                    Cells[c.row].Add(c.column, c);
                }

                
            }

            conn.Close();

     
        }

       

        public string InputOnlyFunction (Cell c)
        {
            return "function calc" + c.cellAddress +
                "(){\r\nreturn Number($(\"#" + c.cellAddress + "\").val());\r\n}";

        }
        public string ForumlaFunction(Cell c)
        {
            List<string> cellRefs = GetCellReferences(c.formula);
            string jsFormula = c.formula+";";

            foreach (string s in cellRefs)
            {
                jsFormula=jsFormula.Replace(s,"calc"+s+"()");
            }

            return "function calc" + c.cellAddress +
                "(){\r\nvar retVal = 0;\r\n retVal " + jsFormula +
                "\r\n$(\"#" + c.cellAddress + "\").text(retVal.toFixed(2));\r\n" +
                "return Number(retVal);\r\n}";

        }

        public string InputFunction(Cell c)
        {

            return "$(\"#"+c.cellAddress+"\").change(function(){ calcFinal();});";

        }

        
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
                string cellRef = key.Replace("$", "");
                retVal.Add(cellRef);

                formula = formula.Replace(key, "");

                match = Regex.Match(formula, regex,
                RegexOptions.IgnoreCase);
            }

            return retVal;

        }

        public HtmlString DynamicScript
        {
            get
            {
                if (Cells == null)
                {
                    return new HtmlString("");
                }
                else
                {

                string calcFinal = "function calcFinal(){";
                string retVal = "";
                foreach (int r in Cells.Keys)
                {
                    foreach (int c in Cells[r].Keys)
                    {                        

                        Cell cell = Cells[r][c];
                        if (cell.cellAddress=="C8")
                        {
                            calcFinal = calcFinal;
                        }

                        if (cell.formula != null)
                        {
                            retVal += ForumlaFunction(cell) + "\r\n";
                            if (!cell.hasDependent)
                            {
                                calcFinal += "calc" + cell.cellAddress + "();\r\n";
                            }
                        }
                        else
                        {
                            if (cell.hasDependent)
                            {
                                retVal += InputOnlyFunction(cell) + "\r\n";
                                retVal += InputFunction(cell) + "\r\n";
                            }
                            else
                            {
                               
                            }
                        }
                    }
                }
                

                

                retVal += SumFunction() +"\r\n";
                retVal += ProductFunction() + "\r\n";

                


                calcFinal += "}\r\n";
                return new HtmlString("$(document).ready(function(){\r\n$(\"#button1\").click(function(){ calcFinal();});" 
                    + retVal + calcFinal + "});"); 
                    }
            
            }
            
        }

        public string SumFunction()
        {
            string retVal = "";
            retVal +="function SUM() {"+"\r\n";
            retVal += "var retVal = arguments[0];" + "\r\n";
            retVal +="for(var i=1; i<arguments.length; i++) " +"\r\n";
            retVal += "{"+"\r\n";
            retVal +="retVal =retVal +arguments[i];" +"\r\n";
            retVal +="}"+"\r\n";
            retVal += "return retVal;" +"\r\n";
            retVal +="}" +"\r\n";
            return retVal;
        }

        public string ProductFunction()
        {
            string retVal = "";
            retVal += "function PRODUCT() {" + "\r\n";
            retVal += "var retVal = arguments[0];" + "\r\n";
            retVal += "for(var i=1; i<arguments.length; i++) " + "\r\n";
            retVal += "{" + "\r\n";
            retVal += "retVal =retVal * arguments[i];" + "\r\n";
            retVal += "}" + "\r\n";
            retVal += "return retVal;" + "\r\n";
            retVal += "}" + "\r\n";
            return retVal;
        }

        public class Cell
        {
            //public string sheetName;
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
            public string FontFamily;
            public double FontSize;
            public string ForeColor;
            public string BackColor;
            public bool Underline;
            public bool Bold;
            public bool Italics;


            public string style
            {
                get 
                {
                    string retval="";
                    if (ForeColor != "000000")
                    {
                        retval+="color:#"+ForeColor+";";
                    }                    
                    if (BackColor != "000000")
                    {
                        retval+="background-color:#"+BackColor+";";
                    }
                    if (Italics)
                    {
                        retval+="font-style:italic;";
                    }
                    if (Bold)
                    {
                        retval+="font-weight:bold;";
                    }
                    if (Underline)
                    {
                        retval+="text-decoration:underline;";
                    }
                    //if (!String.IsNullOrEmpty( FontFamily))
                    //{
                    //    if (FontFamily != "Arial" && FontFamily != "Calibri")
                    //    {
                    //        retval += "font-family:" + FontFamily + ";";
                    //    }
                    //}

                    //if (FontFamily == "Arial" && FontSize == 10)
                    //{
                    //}
                    //else if (FontFamily == "Calibri" && FontSize == 11)
                    //{
                    //}
                    //else
                    //{
                    //    retval += "font-size:" + FontSize + "pt;";
                    //}

                    if (retval != "")
                    {
                        retval = "style="+retval + "}";
                    }
                    
                    return retval;
                }
            }



            //public string FormattedNumber
            //{
            //    get{
            //        decimal d =0;
            //    if (decimal.TryParse(value,out d))
            //    {
            //        return d.ToString(format);
            //    }
            //    else
            //    {
            //        return value;
            //    }
            //    }
            //}
        }
    }
}