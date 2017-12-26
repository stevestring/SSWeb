using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SqlClient;
using System.Data;

namespace SSWeb.Models
{
    public class User
    {
        public List<Worksheet> OwnedWorksheets = new List<Worksheet>();
        public string UserID { get; set; }

        public void GetOwnedWorksheets()
        {

            SqlConnection conn = new SqlConnection(SSWeb.Helpers.Database.ConnectionString);

            conn.Open();

            SqlCommand cmd = new SqlCommand("get_user_worksheets", conn);
            cmd.Parameters.AddWithValue("owner", UserID);
            cmd.CommandType = CommandType.StoredProcedure;

            SqlDataReader dr = cmd.ExecuteReader();

            OwnedWorksheets.Clear();
            while (dr.Read())
            {
                Worksheet w = new Worksheet();
                w.WorksheetId = int.Parse(dr["worksheet_id"].ToString());
                w.Name = dr["name"].ToString();
                w.Description = dr["description"].ToString();
                w.CreatedBy = "Demo";
                w.Owner = dr["owner"].ToString();
                w.Private = bool.Parse(dr["private"].ToString());
                OwnedWorksheets.Add(w);
            }

            conn.Close();
        }

    }    
}