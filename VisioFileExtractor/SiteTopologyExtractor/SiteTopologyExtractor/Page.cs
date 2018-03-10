using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SiteTopologyExtractor
{
    public class Page
    {
        public string Uri { get; set; }
        public string ContentType { get; set; }
        public List<PackagePartItem> PackagePartItems { get; set; }

        public Page()
        {
            PackagePartItems = new List<PackagePartItem>();
        }
    }
}
