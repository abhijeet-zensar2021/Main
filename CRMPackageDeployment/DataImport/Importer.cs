using ExcelDataReader;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRMPackageDeployment.DataImport
{
    public class Importer
    {
        private readonly IOrganizationService _crmservice;
        private readonly string _baseFolder;
        private readonly ITracingService _log;

        public Importer(IOrganizationService crmservice, string baseFolder, ITracingService log)
        {
            _crmservice = crmservice;
            _baseFolder = baseFolder;
            _log = log;
        }

        public void import(ImportConfig config)
        {
            var entities = config.CrmImports
                .Select(i => i.LogicalName)
                .Distinct();

            //Get Metadata
            var requiredMetadata = new Dictionary<string, EntityMetadata>();
            foreach (var ent in entities)
            {
                var entityRequest = new RetrieveEntityRequest()
                {
                    LogicalName = ent,
                    EntityFilters = EntityFilters.Attributes
                };
                var metadata = (RetrieveEntityResponse) _crmservice.Execute(entityRequest);
                if(metadata == null)
                {
                    throw new Exception($"Error: Invalid Logical Name {ent}");
                }
                requiredMetadata.Add(ent, metadata.EntityMetadata);

                foreach (var configItem in config.CrmImports.Where(c => c.ImportType == ImportType.Calender))
                {
                    ProcessCalenderImport(requiredMetadata, configItem);
                }
            }

        }

        private void ProcessCalenderImport (Dictionary<string, EntityMetadata> requiredMetadata, ImportItem configItem)
        {
            using (var fileStream = new FileStream(Path.Combine(_baseFolder, configItem.FileName),FileMode.Open))
            using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(fileStream))
            {
                var headerFactory = new CrmRowHeaderFactory(requiredMetadata[configItem.LogicalName],_crmservice)

            }
        }

    }
}
