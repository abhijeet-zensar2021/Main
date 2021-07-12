using System.IO;
using System.Xml.Serialization;

namespace CRMPackageDeployment.DataImport
{
    public static class ConfigLoader
    {
        public static ImportConfig LoadConfig(string fileName)
        {
            var loader = File.ReadAllText(fileName);
            var xmlSerialiser = new XmlSerializer(typeof(ImportConfig));

            using(TextReader reader = new StringReader(loader))
            {
                return xmlSerialiser.Deserialize(reader) as ImportConfig;
            }

        }
    }
}
