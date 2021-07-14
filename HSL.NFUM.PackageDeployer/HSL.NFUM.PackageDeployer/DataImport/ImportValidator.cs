using ExcelDataReader;
using HSL.NFUM.PackageDeployer.DataImport.AttributeProcessors;
using HSL.NFUM.PackageDeployer.DataImport.AttributeValidators;
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
    public class ImportValidator
    {
        private readonly IOrganizationService _crmService;
        private readonly string _baseFolder;
        private readonly ITraceLogging _log;

        public ImportValidator(IOrganizationService crmService, string baseFolder, ITraceLogging log)
        {
            _crmService = crmService;
            _baseFolder = baseFolder;
            _log = log;
        }

        public bool Validate(ImportConfig config)
        {
            // Lets get the metadata for the entities
            var hasErrored = false;
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

                var metadata = (RetrieveEntityResponse)_crmService.Execute(entityRequest);
                if (metadata == null)
                {
                    throw new Exception($"Error: invalid logical name {ent}");
                }

                requiredMetadata.Add(ent, metadata.EntityMetadata);
            }

            // foreach of the files
            foreach (var configItem in config.CrmImports)
            {
                // create the standard CRM import
                _log.Log($"Validating config item {configItem.Name}", TraceEventType.Information);
                using (var fileStream = new FileStream(Path.Combine(_baseFolder, configItem.FileName), FileMode.Open))
                using (var excelReader = ExcelReaderFactory.CreateOpenXmlReader(fileStream))
                {
                    var headerFactory = new CrmValidatorFactory(requiredMetadata[configItem.LogicalName], _crmService);

                    // generate all the headers, these are important for the parsing
                    int i = 0;
                    var rowProcessors = new List<CrmValidatorBase>();
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
                    var crmStandardImport = new CrmStandardValidator(
                        _crmService,
                        requiredMetadata[configItem.LogicalName],
                        rowProcessors.ToArray(),
                        _log);

                    //go through each of the items and import or update the row
                    var rowNumber = 1;
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
                        
                        _log.Log($"Validating row { rowNumber } RowId {excelReader.GetString(0)}", TraceEventType.Information);
                        var validations = crmStandardImport.ValidateRow(row);
                        if (validations.Length > 0)
                        {
                            foreach (var val in validations)
                            {
                                _log.Log(val, TraceEventType.Error);
                            }
                            hasErrored = true;
                        }

                        rowNumber++;
                    }
                }
            }

            return !hasErrored;
        }
    }
}
