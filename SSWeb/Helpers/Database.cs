using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
namespace SSWeb.Helpers
{
    public static class Database
    {

        public static string ConnectionString
        {
            get
            {
                return ConfigurationManager.ConnectionStrings["ApplicationServices"].ToString();
            }
        }

    }
}