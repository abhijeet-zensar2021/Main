using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace HSL.NFUM.PackageDeployer.DataImport
{
	// configuration of the entities
	public static class ConfigLoader
	{
		public static ImportConfig LoadConfig(string fileName)
		{
			var loader = File.ReadAllText(fileName);
			var xmlSerialiser = new XmlSerializer(typeof(ImportConfig));

			using (TextReader reader = new StringReader(loader))
			{
				return xmlSerialiser.Deserialize(reader) as ImportConfig;
			}
		}
	}
}
