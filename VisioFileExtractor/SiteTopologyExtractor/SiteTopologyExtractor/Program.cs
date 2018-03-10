using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Linq;
using System.IO;
using System.IO.Packaging;
using Newtonsoft.Json;

namespace SiteTopologyExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            PackageItem pkgItem = null;
            //"C:\EGI\Projects\ENABLE\MSDP\Documents\site topology\BUCHAREST ver_20180112.vsdx"
            string fileName = @"C:\EGI\Projects\ENABLE\MSDP\Documents\site topology\BUCHAREST ver_20180112.vsdx";
            //string fileName = @"C:\EGI\Projects\ENABLE\MSDP\Documents\stextractor\SampleFiles\Drawing2.vsdx";

            Package pkg = OpenPackage(fileName);

            if(pkg!=null)
            {
                try
                {
                    pkgItem = IteratePackages(pkg);
                }
                finally
                {
                    pkg.Close();
                }
            }

            string jsonData = JsonConvert.SerializeObject(pkgItem, Newtonsoft.Json.Formatting.Indented);
            File.WriteAllText(Path.GetDirectoryName(fileName) + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".json", jsonData, ASCIIEncoding.UTF8);

            Console.WriteLine("Press <ENTER> to end the execution");
            Console.ReadLine();
        }

        private static PackageItem IteratePackages(Package pkg)
        {
            PackageItem pkgItem = new PackageItem();
            Master master = null;
            PackagePartItem pkgPartItem = null;
            List<MasterMap> masterMaps = new List<MasterMap>();
            Page page = null;

            if (pkg == null)
                return pkgItem;

            foreach (PackagePart pkgPart in pkg.GetParts())
            {
                pkgPartItem = new PackagePartItem();
                Console.WriteLine("=================================================================================");
                Console.WriteLine("Package part URI : {0}", pkgPart.Uri);
                Console.WriteLine("Content Type : {0}", pkgPart.ContentType.ToString());
                pkgPartItem.Uri = pkgPart.Uri.ToString();
                pkgPartItem.ContentType = pkgPart.ContentType;
                pkgItem.PackagePartItems.Add(pkgPartItem);
                try
                {
                    if (pkgPart.ContentType.IndexOf("application/vnd.ms-visio.masters+xml") > -1) //this is page type content type
                    {
                        //masterMap = getMasterMap(pkgPart);
                        masterMaps = getMasterMaps(pkgPart, pkgItem);
                        if (masterMaps != null)
                            pkgItem.MasterMaps = masterMaps;
                    }
                    else if (pkgPart.ContentType.IndexOf("application/vnd.ms-visio.master+xml") > -1) //this is page type content type
                    {
                        master = getMaster(pkgPart, pkgItem);
                        if (master != null)
                            pkgItem.Masters.Add(master);
                    }
                    else if (pkgPart.ContentType.IndexOf("application/vnd.ms-visio.page+xml") > -1) //this is page type content type
                    {
                        page = getPages(pkgPart, pkgItem);
                        if (page != null)
                            pkgItem.Pages.Add(page);
                    }

                    Console.WriteLine("--------------------------- Relationships ---------------------------------------");

                    foreach (var pkgRelationship in pkgPart.GetRelationships())
                    {
                        Console.WriteLine("ID: {0} | Source : {1} -> Target : {2} | Relationship Type : {3}",
                            pkgRelationship.Id, pkgRelationship.SourceUri, pkgRelationship.TargetUri, pkgRelationship.RelationshipType);
                    }
                }
                catch(Exception ex)
                {
                    Console.WriteLine("Error : " + ex.Message);
                }
                finally
                {

                }
                Console.WriteLine("=================================================================================");
            }

            BuildNodeRelationship(pkgItem);

            Console.WriteLine("End of operation");

            return pkgItem;
        }

        private static void BuildNodeRelationship(PackageItem pkgItem)
        {
            foreach (var pagePart in pkgItem.Pages.SelectMany(pg=> pg.PackagePartItems))
            {
                try
                {
                    if (pagePart.Type.Equals("group", StringComparison.OrdinalIgnoreCase) && pagePart.Properties.Count>0 && pagePart.Properties.ContainsKey("Site Code"))
                    {
                        string siteCode = pagePart.Properties["Site Code"];
                        //This is site node
                        var linkedNodes = pkgItem.Pages.SelectMany(pg => pg.PackagePartItems)
                            .Where(ppi => ppi.Type.Equals("shape", StringComparison.OrdinalIgnoreCase)
                            && ((ppi.Properties.ContainsKey("SITE A") && ppi.Properties["SITE A"] == siteCode)
                                || (ppi.Properties.ContainsKey("SITE B") && ppi.Properties["SITE B"] == siteCode)
                                || (ppi.Properties.ContainsKey("Site A") && ppi.Properties["Site A"] == siteCode)
                                || (ppi.Properties.ContainsKey("Site B") && ppi.Properties["Site B"] == siteCode))
                            ); //.FirstOrDefault();
                        if (linkedNodes != null)
                        {
                            foreach (var linkedNode in linkedNodes)
                            {
                                string targetNodeName = null;
                                string siteA = linkedNode.Properties.ContainsKey("SITE A") ? linkedNode.Properties["SITE A"] : linkedNode.Properties["Site A"];
                                string siteB = linkedNode.Properties.ContainsKey("SITE B") ? linkedNode.Properties["SITE B"] : linkedNode.Properties["Site B"];
                                if (siteA == siteCode)
                                    targetNodeName = siteB;
                                else if (siteB == siteCode)
                                    targetNodeName = siteA;

                                if (!string.IsNullOrEmpty(targetNodeName))
                                {
                                    var targetNode = pkgItem.Pages.SelectMany(pg => pg.PackagePartItems)
                                        .Where(ppi => ppi.Type.Equals("group", StringComparison.OrdinalIgnoreCase)
                                        && (ppi.Properties.ContainsKey("Site Code") && ppi.Properties["Site Code"] == targetNodeName)).FirstOrDefault();
                                    if (targetNode != null)
                                    {
                                        pagePart.RelatedNodes.Add("{" + string.Format("'ID': '{0}', 'Name': '{1}', 'Master': '{2}'", targetNode.ID, targetNodeName, targetNode.Master) + "}");
                                    }
                                }
                            }
                        }
                    }
                }
                catch(Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }

        private static Master getMaster(PackagePart pkgPart, PackageItem pkgItem)
        {
            Master master = null;
            XDocument packagePartContent = null;
            Console.WriteLine("--------------------------- Elements ---------------------------------------");
            packagePartContent = getPckagePartContent(pkgPart);
            if (packagePartContent != null)
            {
                XmlReader reader = packagePartContent.CreateReader();
                XmlNamespaceManager namespaceManager = new XmlNamespaceManager(reader.NameTable);
                namespaceManager.AddNamespace("aw", "http://schemas.microsoft.com/office/visio/2012/main");
                namespaceManager.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

                var element = packagePartContent.Descendants().Where(el => el.Name.LocalName.Equals("Shape")).FirstOrDefault();
                //foreach (var element in packagePartContent.Descendants().Where(el => el.Name.LocalName.Equals("Shape")))
                if(element!=null)
                {
                    master = new Master();
                    master.Uri = pkgPart.Uri.ToString();

                    master.Id = pkgPart.Uri.ToString().Replace(@"/visio/masters/master", "").Replace(".xml", "");

                    master.RefId = string.Format("rId{0}", master.Id);

                    if (element.Attribute("Name") != null)
                        master.Name = element.Attribute("Name").Value;
                    if (element.Attribute("Type") != null)
                        master.Type = element.Attribute("Type").Value;

                    var selectedElementCells = element.XPathSelectElements("./aw:Section[@N='Property']/aw:Row/aw:Cell[@N='Label']", namespaceManager);

                    if (selectedElementCells != null)
                    {
                        foreach (var elementCell in selectedElementCells)
                        {
                            if(elementCell.Parent.Name.LocalName.Equals("Row"))
                            {
                                if (!string.IsNullOrEmpty(elementCell.Parent.Attribute("N").Value) &&
                                    !master.Properties.ContainsKey(elementCell.Parent.Attribute("N").Value) && 
                                    !string.IsNullOrEmpty(elementCell.Attribute("V").Value))
                                {
                                    master.Properties.Add(elementCell.Parent.Attribute("N").Value.Trim(), elementCell.Attribute("V").Value.Replace("\n","").Trim());
                                }
                                else
                                {
                                    Console.WriteLine("Duplicate key : " + elementCell.Parent.Attribute("N").Value);
                                }
                            }
                        }
                    }
                }
            }

            return master;
        }

        private static Page getPages(PackagePart pkgPart, PackageItem pkgItem)
        {
            Page page = new Page { Uri = pkgPart.Uri.ToString(), ContentType = pkgPart.ContentType };

            PackagePartItem packagePartItem = null;
            XDocument packagePartContent = null;
            Console.WriteLine("--------------------------- Elements ---------------------------------------");
            packagePartContent = getPckagePartContent(pkgPart);
            if (packagePartContent != null)
            {
                //packagePartItem = new PackagePartItem { Uri = pkgPart.Uri.ToString(), ContentType = pkgPart.ContentType };

                XmlReader reader = packagePartContent.CreateReader();
                XmlNamespaceManager namespaceManager = new XmlNamespaceManager(reader.NameTable);
                namespaceManager.AddNamespace("aw", "http://schemas.microsoft.com/office/visio/2012/main");
                namespaceManager.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

                foreach (var element in packagePartContent.Descendants().Where(el => el.Name.LocalName.Equals("Shape")))
                {
                    try
                    {
                        if (element.Attribute("Name") != null && element.Attribute("Master") != null)
                        {
                            packagePartItem = new PackagePartItem { ContentType = pkgPart.ContentType };
                            Console.WriteLine("\tElement Id: {0} | Type: {1} | Element [Name]: {2} | Master: {3}",
                                element.Attribute("ID").Value, element.Attribute("Type").Value,
                                element.Attribute("Name").Value, element.Attribute("Master").Value);

                            packagePartItem.ID = element.Attribute("ID").Value;
                            packagePartItem.Type = element.Attribute("Type").Value;
                            packagePartItem.Name = element.Attribute("Name").Value;
                            packagePartItem.Master = element.Attribute("Master").Value;

                            MasterMap masterMap = pkgItem.MasterMaps.Where(mm => mm.Id == packagePartItem.Master).FirstOrDefault();
                            if (masterMap != null)
                                packagePartItem.MasterItem = pkgItem.Masters.FirstOrDefault(m => m.RefId == masterMap.RefId);

                            packagePartItem.Uri = string.Format("/visio/masters/master{0}.xml", packagePartItem.Master);

                            //packagePartItem.Properties.Add("ID", element.Attribute("ID").Value);
                            //packagePartItem.Properties.Add("Type", element.Attribute("Type").Value);
                            //packagePartItem.Properties.Add("Name", element.Attribute("Name").Value);
                            //packagePartItem.Properties.Add("Master", element.Attribute("Master").Value);

                            //http://schemas.microsoft.com/office/visio/2012/main
                            var selectedElements = element.XPathSelectElements("./aw:Section[@N='Property']/aw:Row/aw:Cell", namespaceManager);
                            //int indx = 0;
                            string parkedValue = null;
                            foreach (var selectedElement in selectedElements)
                            {
                                bool isKeyFound = true;
                                if (selectedElement.Attribute("N") != null && 
                                    (selectedElement.Attribute("N").Value.Equals("Value") || 
                                        selectedElement.Attribute("N").Value.Equals("Label")))
                                {
                                    string key = selectedElement.Parent.FirstAttribute.Value;
                                    if (packagePartItem.MasterItem != null && packagePartItem.MasterItem.Properties.ContainsKey(key))
                                        key = packagePartItem.MasterItem.Properties[key];
                                        //key += "_" + packagePartItem.MasterItem.Properties[key];
                                    else
                                    {
                                        isKeyFound = false;
                                        Console.WriteLine("Key Not found " + key);

                                        if (selectedElement.Attribute("N").Value.Equals("Label"))
                                        {
                                            key = selectedElement.Attribute("V").Value;
                                            isKeyFound = true;
                                        }
                                    }

                                    if (selectedElement.Attribute("V") != null)
                                    {
                                        Console.WriteLine("\t\tCell Value : {0}", selectedElement.Attribute("V").Value);

                                        if (selectedElement.Attribute("N").Value.Equals("Value"))
                                            parkedValue = selectedElement.Attribute("V").Value.Replace("\n", "");

                                        if (isKeyFound && parkedValue!=null)
                                        {
                                            if(!packagePartItem.Properties.ContainsKey(key))
                                                packagePartItem.Properties.Add(key, parkedValue);
                                            parkedValue = null;
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("\t\tCell Value : {0}", selectedElement.Value);
                                        if (selectedElement.Attribute("N").Value.Equals("Value"))
                                            parkedValue = selectedElement.Value.Replace("\n", "");

                                        if (isKeyFound && parkedValue != null)
                                        {
                                            if(!packagePartItem.Properties.ContainsKey(key))
                                                packagePartItem.Properties.Add(key, parkedValue);
                                            parkedValue = null;
                                        }
                                    }
                                }
                            }

                            page.PackagePartItems.Add(packagePartItem);
                        }
                    }
                    catch(Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
            Console.WriteLine("--------------------------- End Of Elements ---------------------------------------");

            return page;
        }

        private static List<MasterMap> getMasterMaps(PackagePart pkgPart, PackageItem pkgItem)
        {
            List<MasterMap> masterMaps = new List<MasterMap>();
            MasterMap masterMap = null;
            XDocument packagePartContent = null;
            Console.WriteLine("--------------------------- Elements ---------------------------------------");
            packagePartContent = getPckagePartContent(pkgPart);
            if (packagePartContent != null)
            {
                /*XmlReader reader = packagePartContent.CreateReader();
                XmlNamespaceManager namespaceManager = new XmlNamespaceManager(reader.NameTable);
                namespaceManager.AddNamespace("aw", "http://schemas.microsoft.com/office/visio/2012/main");
                namespaceManager.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");*/

                foreach (var element in packagePartContent.Descendants().Where(el => el.Name.LocalName.Equals("Master")))
                {
                    masterMap = new MasterMap();
                    if (element.Attribute("ID") != null)
                        masterMap.Id = element.Attribute("ID").Value;
                    if (element.Attribute("Name") != null)
                        masterMap.Name = element.Attribute("Name").Value;
                    if (element.Attribute("UniqueID") != null)
                        masterMap.UniqueId = element.Attribute("UniqueID").Value;
                    var relElement = element.Descendants().Where(el => el.Name.LocalName.Equals("Rel")).FirstOrDefault();
                    if(relElement!=null)
                    {
                        masterMap.RefId = relElement.FirstAttribute.Value.Trim();
                    }

                    if (masterMap != null)
                        masterMaps.Add(masterMap);
                }
            }

            return masterMaps;
        }

        private static XDocument getPckagePartContent(PackagePart pkgPart)
        {
            XDocument document = null;

            document = XDocument.Load(pkgPart.GetStream());

            return document;
        }

        private static Package OpenPackage(string filePath)
        {
            Package visioFile = null;

            try
            {
                if (File.Exists(filePath))
                    visioFile = Package.Open(filePath, FileMode.Open, FileAccess.Read);
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error:: {0}", ex.Message);
            }

            return visioFile;
        }
    }
}
