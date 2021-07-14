using ExcelDataReader;
using HSL.NFUM.PackageDeployer.DataImport.AttributeProcessors;
using HSL.NFUM.PackageDeployer.DataImport.Import;
using HSL.NFUM.PackageDeployer.DataImport.Interfaces;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace HSL.NFUM.PackageDeployer.DataImport
{

    public class Importer
    {
        private readonly IOrganizationService _crmService;
        private readonly string _baseFolder;
        private readonly ITraceLogging _log;

        public Importer(IOrganizationService crmService, string baseFolder, ITraceLogging log)
        {
            _crmService = crmService;
            _baseFolder = baseFolder;
            _log = log;
        }

        public void Import(ImportConfig config)
        {

            // Lets get the metadata for the entities
            var entities = config.CrmImports
                .Select(i => i.LogicalName)
                .Distinct();

            // Get all the requird metadata for the imports
            var requiredMetadata = new Dictionary<string, EntityMetadata>();
            foreach (var ent in entities)
            {
                var entityRequest = new RetrieveEntityRequest()
                {
                    LogicalName = ent,
                    EntityFilters = EntityFilters.Attributes
                };

                var metadata = (RetrieveEntityResponse) _crmService.Execute(entityRequest);
                if (metadata == null)
                {
                    throw new Exception($"Error: invalid logical name {ent}");
                }

                requiredMetadata.Add(ent, metadata.EntityMetadata);
            }

            // foreach of the files
            foreach (var configItem in config.CrmImports.Where(c => c.ImportType == ImportType.Standard))
            {
                ProcessStandardImport(requiredMetadata, configItem);
            }

            // Document Template one, we can refactor this later
            foreach (var configItem in config.CrmImports.Where(c => c.ImportType == ImportType.DocumentTemplate))
            {
                ProcessDataTemplateImport(requiredMetadata, configItem);
            }

            // calendar one
            foreach (var configItem in config.CrmImports.Where(c => c.ImportType == ImportType.Calendar))
            {
                ProcessCalendarImport(requiredMetadata, configItem);
            }
        }

        private void ProcessCalendarImport(Dictionary<string, EntityMetadata> requiredMetadata, ImportItem configItem)
        {
            // create the standard CRM import
            _log.Log($"Processing config item {configItem.Name}", TraceEventType.Information);
            using (var fileStream = new FileStream(Path.Combine(_baseFolder, configItem.FileName), FileMode.Open))
            using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(fileStream))
            {
                var headerFactory = new CrmRowHeaderFactory(requiredMetadata[configItem.LogicalName], _crmService);

                // generate all the headers, these are important for the parsing
                int i = 0;
                var rowProcessors = new List<CrmRowProcessor>();
                excelReader.Read();
                while (excelReader.FieldCount > i)
                {
                    var headerLogicalName = excelReader.GetString(i);

                    // break of out while loops after we've finished with the headers
                    if (string.IsNullOrEmpty(headerLogicalName))
                    {
                        break;
                    }

                    else
                    {
                        rowProcessors.Add(headerFactory.GenerateHeader(headerLogicalName));
                    }

                    i++;
                }

                // generate the entities to be read in
                var crmCalendarImport = new CrmCalendarImport(
                    _crmService,
                    requiredMetadata[configItem.LogicalName],
                    rowProcessors.ToArray(),
                    _log);

                //go through each of the items and import or update the row
                var allRows = new List<string[]>();
                while (excelReader.Read())
                {
                    // since id is always first if empty it must have process everything
                    if (string.IsNullOrEmpty(excelReader.GetString(0)))
                    {
                        break;
                    }

                    var row = new string[rowProcessors.Count()];
                    foreach (var c in Enumerable.Range(0, rowProcessors.Count()))
                    {
                        row[c] = excelReader.GetValue(c)?.ToString();
                    }

                    allRows.Add(row);
                }
                crmCalendarImport.UpsertRows(allRows);
            }
        }

        private void ProcessDataTemplateImport(Dictionary<string, EntityMetadata> requiredMetadata, ImportItem configItem)
        {
            // create the standard CRM import
            _log.Log($"Processing config item {configItem.Name}", TraceEventType.Information);
            using (var fileStream = new FileStream(Path.Combine(_baseFolder, configItem.FileName), FileMode.Open))
            using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(fileStream))
            {
                var headerFactory = new CrmRowHeaderFactory(requiredMetadata[configItem.LogicalName], _crmService);

                // generate all the headers, these are important for the parsing
                int i = 0;
                var rowProcessors = new List<CrmRowProcessor>();
                excelReader.Read();
                while (excelReader.FieldCount > i)
                {
                    var headerLogicalName = excelReader.GetString(i);

                    // break of out while loops after we've finished with the headers
                    if (string.IsNullOrEmpty(headerLogicalName))
                    {
                        break;
                    }

                    if (headerLogicalName.Equals("Content"))
                    {
                        var pathForBinaries = Path.Combine(_baseFolder, configItem.FileName);
                        rowProcessors.Add(headerFactory.GenerateContentHeader(Path.GetDirectoryName(pathForBinaries) + "_binaries"));
                    }
                    else
                    {
                        rowProcessors.Add(headerFactory.GenerateHeader(headerLogicalName));
                    }

                    i++;
                }

                // generate the entities to be read in
                var crmStandardImport = new CrmStandardImport(
                    _crmService,
                    requiredMetadata[configItem.LogicalName],
                    rowProcessors.ToArray(),
                    _log);

                //go through each of the items and import or update the row
                while (excelReader.Read())
                {
                    // since id is always first if empty it must have process everything
                    if (string.IsNullOrEmpty(excelReader.GetString(0)))
                    {
                        break;
                    }

                    var row = new string[rowProcessors.Count];
                    foreach (var c in Enumerable.Range(0, rowProcessors.Count))
                    {
                        row[c] = excelReader.GetValue(c)?.ToString();
                    }

                    crmStandardImport.UpsertRow(row);
                }
            }
        }

        private void ProcessStandardImport(Dictionary<string, EntityMetadata> requiredMetadata, ImportItem configItem)
        {
            // create the standard CRM import
            _log.Log($"Processing config item {configItem.Name}", TraceEventType.Information);
            using (var fileStream = new FileStream(Path.Combine(_baseFolder, configItem.FileName), FileMode.Open))
            using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(fileStream))
            {
                var headerFactory = new CrmRowHeaderFactory(requiredMetadata[configItem.LogicalName], _crmService);

                // generate all the headers, these are important for the parsing
                int i = 0;
                var rowProcessors = new List<CrmRowProcessor>();
                excelReader.Read();
                while (excelReader.FieldCount > i)
                {
                    var headerLogicalName = excelReader.GetString(i);

                    // break of out while loops after we've finished with the headers
                    _log.Log($"Processing header {headerLogicalName}", TraceEventType.Information);
                    if (string.IsNullOrEmpty(headerLogicalName))
                    {
                        break;
                    }

                    rowProcessors.Add(headerFactory.GenerateHeader(headerLogicalName));
                    i++;
                }

                // generate the entities to be read in
                var crmStandardImport = new CrmStandardImport(
                    _crmService,
                    requiredMetadata[configItem.LogicalName],
                    rowProcessors.ToArray(),
                    _log);

                //go through each of the items and import or update the row
                while (excelReader.Read())
                {
                    var row = new string[rowProcessors.Count()];

                    // since id is always first if empty it must have process everything
                    if (string.IsNullOrEmpty(excelReader.GetString(0)))
                    {
                        break;
                    }

                    foreach (var c in Enumerable.Range(0, rowProcessors.Count()))
                    {
                        row[c] = excelReader.GetValue(c)?.ToString();
                    }

                    crmStandardImport.UpsertRow(row);
                }
            }
        }
    }
}
