using System.Xml.Serialization;

namespace CRMPackageDeployment.DataImport
{
    public class ImportItem
    {
        [XmlElement("Name")]
        public string Name { get; set; }

        [XmlElement("FileName")]
        public string FileName { get; set; }

        [XmlElement("LogicalName")]
        public string LogicalName { get; set; }

        [XmlElement("ImportType")]
        public ImportType ImportType { get; set; }
    }
}
