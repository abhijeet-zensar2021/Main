using System.Collections.Generic;
using System.Xml.Serialization;

namespace CRMPackageDeployment.DataImport
{
    [XmlRoot("ImportConfg")]
    public class ImportConfig
    {
      public List<ImportItem> CrmImports { get; set; }
    }
}
