using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Linq;

namespace HSL.NFUM.PackageDeployer.DataImport.AttributeProcessors
{
    public class CrmRowHeaderFactory
    {
        private readonly EntityMetadata _metadata;
        private readonly IOrganizationService _crmService;

        public CrmRowHeaderFactory(EntityMetadata metadata, IOrganizationService organizationService)
        {
            _metadata = metadata;
            _crmService = organizationService;
        }

        public CrmRowProcessor GenerateHeader(string displayName)
        {
            // find the matching item in the metadata
            var attr = _metadata.Attributes.SingleOrDefault(a => a?.DisplayName?.UserLocalizedLabel?.Label.Equals(displayName) == true);
            if (attr == null)
            {
                throw new Exception($"Invalid header: {displayName} for {_metadata.LogicalName}");
            }

            if (attr.AttributeType.Value == AttributeTypeCode.String)
            {
                return new StringProcessor(attr.LogicalName, attr);
            }
            if (attr.AttributeType.Value == AttributeTypeCode.Uniqueidentifier)
            {
                return new UniqueIdentifierProcessor(attr.LogicalName, attr);
            }
            if (attr.AttributeType.Value == AttributeTypeCode.Picklist)
            {
                return new OptionsetProcessor(attr.LogicalName, attr);
            }
            if (attr.AttributeType.Value == AttributeTypeCode.Status)
            {
                return new StatusProcessor(attr.LogicalName, attr);
            }
            if (attr.AttributeType.Value == AttributeTypeCode.Boolean)
            {
                return new BooleanProcessor(attr.LogicalName, attr);
            }
            if (attr.AttributeType.Value == AttributeTypeCode.Lookup)
            {
                return new LookupProcessor(attr.LogicalName, attr, _crmService);
            }
            if (attr.AttributeType.Value == AttributeTypeCode.EntityName)
            {
                return new EntityNameProcessor(attr.LogicalName, attr);
            }
            if (attr.AttributeType.Value == AttributeTypeCode.Memo)
            {
                return new MemoProcessor(attr.LogicalName, attr);
            }
            if (attr.AttributeType.Value == AttributeTypeCode.DateTime)
            {
                return new DateTimeProcessor(attr.LogicalName, attr);
            }
            if (attr.AttributeType.Value == AttributeTypeCode.Integer)
            {
                return new IntegerProcessor(attr.LogicalName, attr);
            }

            throw new NotImplementedException($"datatype {attr.AttributeTypeName.Value} of type {attr.AttributeType.Value} for {attr.LogicalName} has not been implemented");
        }

        public CrmRowProcessor GenerateContentHeader(string contentBaseFolder)
        {
            var entitiesReq = new RetrieveAllEntitiesRequest()
            {
                 EntityFilters = EntityFilters.Entity
            };
            var res = (RetrieveAllEntitiesResponse)_crmService.Execute(entitiesReq);
            var map = res
                .EntityMetadata
                .Where(x => x.ObjectTypeCode.HasValue)
                .ToDictionary(x => x.LogicalName, x => x.ObjectTypeCode.Value);

            return new FileContentProcessor(contentBaseFolder, map);
        }
    }
}
