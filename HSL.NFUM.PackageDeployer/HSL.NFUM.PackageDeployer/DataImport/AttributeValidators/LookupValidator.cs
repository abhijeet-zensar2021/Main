using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;

namespace HSL.NFUM.PackageDeployer.DataImport.AttributeValidators
{

    public class LookupValidator : CrmValidatorBase
    {
        private readonly IOrganizationService _crmService;
        private readonly Dictionary<string, Dictionary<string, Entity>> _cache = new Dictionary<string, Dictionary<string, Entity>>();

        private LookupAttributeMetadata LookupAttribute { get { return (LookupAttributeMetadata)this.Attribute; } }

        public LookupValidator(string logicalName, AttributeMetadata attr, IOrganizationService crmService) : base(logicalName, attr)
        {
            _crmService = crmService;
        }

        public override string[] Validate(string val)
        {
            if (!val.Contains(":"))
            {
                return new[] { $"Error: {Attribute.EntityLogicalName}.{Attribute.LogicalName} needs to be in format lookup:value" };
            }

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

                var results = _crmService.RetrieveMultiple(query).Entities
                    .ToDictionary(r => r.GetAttributeValue<string>(field));
                _cache.Add(field, results);
            }

            var cache = _cache[field];
            if (!cache.ContainsKey(value))
            {
                return new[] { $"Error: Cannot Find {entityName}.{field} with value {value}" };
            }

            return new string[] { };
        }
    }
}
