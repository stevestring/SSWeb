using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace SSWeb.Helpers
{
    public class JSONHelper
    {

        public static T Deserialise<T>(string json)
        {
            T obj = Activator.CreateInstance<T>();

            MemoryStream ms = new MemoryStream(Encoding.Unicode.GetBytes(json));

            DataContractJsonSerializer serializer = new DataContractJsonSerializer(obj.GetType());
            obj = (T)serializer.ReadObject(ms); // <== Your missing line

            return obj;

        }

    }
}