using DocumentFormat.OpenXml.Packaging;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace HSL.NFUM.PackageDeployer.DataImport.AttributeProcessors
{

    // special class that is used to populate the CrmRowProcess
    public class FileContentProcessor : CrmRowProcessor
    {
        private readonly string _baseFolder;
        private readonly Dictionary<string, int> _entityMap;

        public FileContentProcessor(string baseFolder, Dictionary<string, int> entityMap) : base("content", null)
        {
            _baseFolder = baseFolder;
            _entityMap = entityMap;
        }

        public override void AddValue(ref Entity record, string filename)
        {
            var fileContent = File.ReadAllBytes(Path.Combine(_baseFolder, filename));

            var entityDetails = GetETNAndOTCFromDocument(fileContent);
            Console.WriteLine(entityDetails?.EntityTypeName);
            var item = entityDetails != null ?
                UpdateTemplate(fileContent, entityDetails).FileContents :
                fileContent;

            record[this.LogicalName] = Convert.ToBase64String(item);
        }

        /// <summary>
        /// Reads in the files, updates them to the correct string
        /// Writes them back out
        /// </summary>
        /// <param name="filePath"></param>
        DocumentUpdate UpdateTemplate(byte[] content, EntityDetails entityDetails)
        {
            if (!_entityMap.ContainsKey(entityDetails.EntityTypeName))
            {
                throw new Exception($"Error: target environment does not contain entity {entityDetails.EntityTypeName}");
            }
            var newNumber = _entityMap[entityDetails.EntityTypeName];

            return new DocumentUpdate(entityDetails.EntityTypeName, entityDetails.ObjectTypeCode, newNumber, content);
        }

        private EntityDetails GetETNAndOTCFromDocument(byte[] content)
        {
            const string rootPath = @"urn:microsoft-crm/document-template/";

            string xmlBody = null;
            try
            {
                // open the file and read the contents 
                using (var ms = new MemoryStream(content))
                {
                    // create a word doc from stream
                    WordprocessingDocument wordDoc = WordprocessingDocument.Open(ms, true);
                    OpenXmlPart mainDocPart = wordDoc.MainDocumentPart;

                    var custParts = wordDoc
                        .MainDocumentPart
                        .Parts
                        .Where(p => p.OpenXmlPart is CustomXmlPart)
                        .Select(c => (c.OpenXmlPart as CustomXmlPart).CustomXmlPropertiesPart);

                    foreach (var part in custParts)
                    {
                        if (part.RootElement.InnerXml.Contains(rootPath))
                        {
                            xmlBody = part.RootElement.InnerXml;
                            break;
                        }
                    }
                }

                if (xmlBody != null)
                {
                    // get the ETC from the first instance we can find of the ds:uri attribute.
                    var ns = @"{http://schemas.openxmlformats.org/officeDocument/2006/customXml}";
                    using (var textStream = new StringReader(xmlBody))
                    {
                        XDocument doc = XDocument.Load(textStream);
                        XAttribute attrib = doc
                            .Descendants($"{ns}schemaRef")
                            .Attributes($"{ns}uri")
                            .Where<XAttribute>(a => a.Value.StartsWith(rootPath))
                            .FirstOrDefault();

                        var etn_otc = attrib.Value.Substring(rootPath.Length).Split('/');

                        return new EntityDetails
                        {
                            EntityTypeName = etn_otc[0],
                            ObjectTypeCode = int.Parse(etn_otc[1])
                        };
                    }
                }
            }
            catch (NullReferenceException) { }

            // don't need to do an update
            return null;
        }
    }

    public class EntityDetails
    {
        public string EntityTypeName { get; set; }
        public int ObjectTypeCode { get; set; }
    }

    public class DocumentUpdate
    {
        private string EntityTypeName { get; }
        private int ObjectTypeCode { get; }
        private int TargetOTC { get; }
        public byte[] FileContents { get; set;  }

        public DocumentUpdate(string entityTypeName, int objectTypeCode, int targetOTC, byte[] fileContents)
        {
            EntityTypeName = entityTypeName;
            ObjectTypeCode = objectTypeCode;
            TargetOTC = targetOTC;
            FileContents = fileContents;

            UpdateOTCCodes();
        }

        private void UpdateOTCCodes()
        {
            using (var ms = new MemoryStream(this.FileContents))
            {
                // create a word doc from stream
                WordprocessingDocument wordDoc = WordprocessingDocument.Open(ms, true);

                // instantiate the parts of the word doc
                OpenXmlPart mainDocPart = wordDoc.MainDocumentPart;
                IEnumerable<OpenXmlPart> docHeaderParts = mainDocPart.Parts.Where(p => p.OpenXmlPart is HeaderPart).Select(p => p.OpenXmlPart);
                IEnumerable<OpenXmlPart> docFooterParts = mainDocPart.Parts.Where(p => p.OpenXmlPart is FooterPart).Select(p => p.OpenXmlPart);
                IEnumerable<OpenXmlPart> customParts = mainDocPart.Parts.Where(p => p.OpenXmlPart is CustomXmlPart).Select(p => p.OpenXmlPart);

                IEnumerable<OpenXmlPart> customPropParts =
                    from parent in customParts
                    from child in parent.Parts
                    where child.OpenXmlPart is CustomXmlPropertiesPart
                    select child.OpenXmlPart;

                // change type codes in each part
                UpdateDocumentPart(mainDocPart, EntityTypeName, ObjectTypeCode, TargetOTC);

                UpdateDocumentParts(docHeaderParts, EntityTypeName, ObjectTypeCode, TargetOTC);
                UpdateDocumentParts(docFooterParts, EntityTypeName, ObjectTypeCode, TargetOTC);
                UpdateDocumentParts(customParts, EntityTypeName, ObjectTypeCode, TargetOTC);
                UpdateDocumentParts(customPropParts, EntityTypeName, ObjectTypeCode, TargetOTC);

                // get wordDoc back into format required for CRM
                wordDoc.Close();

                this.FileContents = ms.ToArray();
            };
        }

        /// <summary>
        /// swaps out bad entity type codes for good ones in the given Word Doc part
        /// </summary>
        /// <param name="part">the Open Xml Part needing to be changed</param>
        /// <param name="sourceOTC">source type code needing to be changed</param>
        /// <param name="targetOTC">destination type code</param>
        private void UpdateDocumentPart(OpenXmlPart part, string sourceETC, int sourceOTC, int targetOTC)
        {
            string docText;
            using (StreamReader sr = new StreamReader(part.GetStream()))
            {
                docText = sr.ReadToEnd();
            }

            docText = docText.Replace($"{sourceETC}/{sourceOTC}/", $"{sourceETC}/{targetOTC}/");

            using (StreamWriter sw = new StreamWriter(part.GetStream()))
            {
                sw.Write(docText);
                sw.Flush();
            }
        }

        /// <summary>
        /// changes all bad entity type codes OpenXmlParts in pair to good code
        /// </summary>
        /// <param name="pair">container of desired parts</param>
        /// <param name="sourceOTC">source type code needing to be changed</param>
        /// <param name="goodTC">destination type code</param>
        private void UpdateDocumentParts(IEnumerable<OpenXmlPart> pair, string sourceETC, int sourceOTC, int targetOTC)
        {
            foreach (OpenXmlPart part in pair)
            {
                UpdateDocumentPart(part, sourceETC, sourceOTC, targetOTC);
            }
        }
    }
}
