using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SiteTopologyExtractor
{
    public class PackageItem
    {
        public string Uri { get; set; }
        public string ContentType { get; set; }
        public List<Page> Pages { get; set; }
        public List<Master> Masters { get; set; }
        public List<MasterMap> MasterMaps { get; set; }
        public List<PackagePartItem> PackagePartItems { get; set; }

        public PackageItem()
        {
            Pages = new List<Page>();
            Masters = new List<Master>();
            MasterMaps = new List<MasterMap>();
            PackagePartItems = new List<PackagePartItem>();
        }
    }

    public class PackagePartItem
    {
        public string ID { get; set; }
        public string Type { get; set; }
        public string Name { get; set; }
        public string Master { get; set; }
        public string Uri { get; set; }
        public string ContentType { get; set; }
        public Dictionary<string, string> Properties { get; set; }
        public Master MasterItem { get; set; }
        public List<String> RelatedNodes { get; set; }

        public PackagePartItem()
        {
            Properties = new Dictionary<string, string>();
            RelatedNodes = new List<string>();
        }
    }
}
