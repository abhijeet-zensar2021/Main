using CRMPackageDeployment.DataImport.AttributeProcessor;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Linq;

namespace CRMPackageDeployment.Factories
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
            var attr = _metadata.Attributes.SingleOrDefault(a => a?.DisplayName?.UserLocalizedLabel?.Label.Equals(displayName) == true);
            if(attr == null)
            {
                throw new Exception($"Invalid header: {displayName} for {_metadata.LogicalName}");
            }
            if(attr.AttributeType.Value == AttributeTypeCode.String)
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
                return new LookupProcessor(attr.LogicalName, attr);
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
            throw new NotImplementedException();
        }
    }
}
