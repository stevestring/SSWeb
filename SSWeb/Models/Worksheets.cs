using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
namespace SSWeb.Models
{
    public class Worksheets:List<Models.Worksheet>
    {

        public void GetWorkSheets()
        {

            SqlConnection conn = new SqlConnection(SSWeb.Helpers.Database.ConnectionString);

            conn.Open();

            SqlCommand cmd = new SqlCommand("get_worksheets", conn);
            cmd.CommandType = CommandType.StoredProcedure;

            SqlDataReader dr = cmd.ExecuteReader();

            this.Clear();
            while (dr.Read())
            {
                Worksheet w = new Worksheet();
                w.WorksheetId = int.Parse(dr["worksheet_id"].ToString());
                w.Name = dr["name"].ToString();
                if (String.IsNullOrEmpty(w.Name))
                {
                    w.Name = "NA";
                }
                w.Description = dr["description"].ToString();
                //w.CreatedBy = "Demo";
                w.Owner = dr["owner"].ToString();
                w.Private = bool.Parse(dr["private"].ToString());
                this.Add(w);
            }

            conn.Close();


        }
    }
}