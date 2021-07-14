using System.Collections.Generic;
using System.Xml.Serialization;

namespace HSL.NFUM.PackageDeployer.DataImport
{
    // Define other methods and classes here
    [XmlRoot("ImportConfig")]
	public class ImportConfig
	{
		public List<ImportItem> CrmImports { get; set; }
	}
}
