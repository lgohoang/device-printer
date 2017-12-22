using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace CLI
{
    [DataContract]
    class DeviceStatus
    {
        [DataMember(Name = "code")]
        public int Code;

        [DataMember(Name = "name")]
        public string Name;

        [DataMember(Name = "source")]
        public string Source;

        [DataMember(Name = "description")]
        public string Description;

        // importance level
        private static DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(DeviceStatus));
        public virtual string Serialize()
        {
            //Create a stream to serialize the object to.  
            MemoryStream ms = new MemoryStream();
            serializer.WriteObject(ms, this);
            byte[] json = ms.ToArray();
            ms.Close();
            return Encoding.UTF8.GetString(json, 0, json.Length);
        }
    }
}
