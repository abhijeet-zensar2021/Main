using HSL.NFUM.PackageDeployer.DataImport.Extensions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;

namespace HSL.NFUM.PackageDeployer.DataImport.AttributeProcessors
{

    public class LookupProcessor : CrmRowProcessor
    {
        private readonly IOrganizationService _crmService;
        private readonly Dictionary<string, Lookup<string, Entity>> _cache = new Dictionary<string, Lookup<string, Entity>>();

        private LookupAttributeMetadata LookupAttribute { get { return (LookupAttributeMetadata)this.Attribute; } }

        public LookupProcessor(string logicalName, AttributeMetadata attr, IOrganizationService crmService) : base(logicalName, attr)
        {
            _crmService = crmService;
        }

        public override void AddValue(ref Entity record, string val)
        {
            // split the action by : to work out the key and the matching value
            var field = val.Split(':')[0];
            var value = val.Split(':')[1];
            var entityName = LookupAttribute.Targets.First();

            // do a retrieve on the entity
            if (!_cache.ContainsKey(field))
            {
                var query = new QueryExpression(entityName)
                {
                    ColumnSet = new ColumnSet(field)
                };
                query.Criteria.AddCondition(field, ConditionOperator.NotNull);

                var results = (Lookup<string, Entity>) _crmService
                    .RetrieveAllRecords(query)
                    .ToLookup(x => x.GetAttributeValue<string>(field));

                _cache.Add(field, results);
            }

            var cache = _cache[field][value];
            if (!cache.Any())
            {
                throw new Exception($"Error: Cannot Find {entityName}.{field} with value {value}");
            }
            if (cache.Count() > 1)
            {
                throw new Exception($"Error: Multiple {entityName}.{field} with value {value}");
            }

            record[this.LogicalName] = cache.First().ToEntityReference();
        }
    }
}
