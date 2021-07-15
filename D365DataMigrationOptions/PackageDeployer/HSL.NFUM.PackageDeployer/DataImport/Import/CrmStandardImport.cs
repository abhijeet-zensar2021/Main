using HSL.NFUM.PackageDeployer.DataImport.AttributeProcessors;
using HSL.NFUM.PackageDeployer.DataImport.Extensions;
using HSL.NFUM.PackageDeployer.DataImport.Interfaces;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace HSL.NFUM.PackageDeployer.DataImport.Import
{
    public class CrmStandardImport
    {
        private readonly IOrganizationService _crmService;
        private readonly EntityMetadata _metadata;
        private readonly SortedDictionary<Guid, Entity> _allRecords;
        private readonly CrmRowProcessor[] _rowProcessors;
        private readonly ITraceLogging _log;

        public CrmStandardImport(IOrganizationService crmService, EntityMetadata metadata, CrmRowProcessor[] rowProcessors, ITraceLogging log)
        {
            _crmService = crmService;
            _metadata = metadata;
            _rowProcessors = rowProcessors;
            _log = log;

            // retrieve all the of the record that this has so we can compare
            // use sorted dictionary for quick retrieval
            var columns = new ColumnSet(rowProcessors.SelectMany(x => x.GetColumnsetColumns()).ToArray());
            _log.Log($"Retrieving all {metadata.LogicalName} with columns {string.Join(",", columns.Columns)}", TraceEventType.Information);

            var allEntityQuery = new QueryExpression(metadata.LogicalName)
            {
                ColumnSet = columns,
            };
            var records = crmService.RetrieveAllRecords(allEntityQuery).ToDictionary(x => x.Id);
            _allRecords = new SortedDictionary<Guid, Entity>(records);
        }

        internal void UpsertRow(string[] row)
        {
            // create and populate the rquired entity
            var ent = new Entity(_metadata.LogicalName);
            for (var i = 0; i < _rowProcessors.Count(); i++)
            {
                _rowProcessors[i].AddValue(ref ent, row[i]);
                if (_rowProcessors[i].LogicalName.Equals(_metadata.PrimaryIdAttribute, StringComparison.OrdinalIgnoreCase))
                {
                    ent.Id = (Guid)ent[_rowProcessors[i].LogicalName];
                }
            }

            if (!_allRecords.ContainsKey(ent.Id))
            {
                _log.Log($"Creating : {_metadata.LogicalName} : {ent.Id}", TraceEventType.Information);
                _crmService.Create(ent);
            }
            else
            {
                // update if different
                var currentRecord = _allRecords[ent.Id];
                var differingAttributes = _rowProcessors
                    .Where(rp =>
                        (!(ent[rp.LogicalName] is DateTime) && !Equals(ent[rp.LogicalName], (currentRecord.Contains(rp.LogicalName) ? currentRecord[rp.LogicalName] : null))) ||
                        ((ent[rp.LogicalName] is DateTime) && !Equals(ent.GetAttributeValue<DateTime>(rp.LogicalName).ToUniversalTime(), (currentRecord.Contains(rp.LogicalName) ? currentRecord[rp.LogicalName] : null))) // when comparing datetime, use universal
                        )
                    .ToList();

                if (differingAttributes.Any())
                {
                    _log.Log($"Updating : {_metadata.LogicalName} : {ent.Id} - Updated: {String.Join(",",differingAttributes.Select(x => x?.LogicalName))}", TraceEventType.Information);
                    _crmService.Update(ent);
                }
                else
                {
                    // no change for the one that is in the system
                    _log.Log($"No Change : {_metadata.LogicalName} : {ent.Id}", TraceEventType.Information);
                }
            }
        }
    }
}
