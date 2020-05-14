using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace CodeStack
{
    [DataContract]
    public class SnapshotInfo
    {
        [DataMember]
        public int Revision { get; set; }
        
        [DataMember]
        public DateTime TimeStamp { get; set; }
    }

    [DataContract]
    public class PropertiesSnapshot
    {
        [DataMember]
        public Dictionary<string, string> Properties { get; set; }
    }
}
