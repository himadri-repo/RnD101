using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SiteTopologyExtractor
{
    public class Master
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
        public string Uri { get; set; }
        public string RefId { get; set; }
        public Dictionary<string,string> Properties { get; set; }

        public Master()
        {
            Properties = new Dictionary<string, string>();
        }
    }

    public class MasterMap
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string UniqueId  { get; set; }
        public string RefId { get; set; }
        public string Target { get; set; }
    }
}
