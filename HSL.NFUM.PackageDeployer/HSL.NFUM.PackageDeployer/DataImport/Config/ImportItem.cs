using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace HSL.NFUM.PackageDeployer.DataImport
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
