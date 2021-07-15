using HSL.NFUM.PackageDeployer.DataImport.AttributeValidators;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Linq;

namespace HSL.NFUM.PackageDeployer.DataImport.AttributeProcessors
{
    public class CrmValidatorFactory
    {
        private readonly EntityMetadata _metadata;
		private readonly IOrganizationService _crmService;

        public CrmValidatorFactory(EntityMetadata metadata, IOrganizationService organizationService)
        {
            _metadata = metadata;
			_crmService = organizationService;
        }

		public CrmValidatorBase GenerateHeader(string displayName)
		{
			// find the matching item in the metadata
			var attr = _metadata.Attributes.SingleOrDefault(a => a?.DisplayName?.UserLocalizedLabel?.Label.Equals(displayName) == true);
			if (attr == null)
			{
				throw new Exception($"Invalid header: {displayName} for {_metadata.LogicalName}");
			}

			if (attr.AttributeType.Value == AttributeTypeCode.String)
			{
				return new StringValidator(attr.LogicalName, attr);
			}
            if (attr.AttributeType.Value == AttributeTypeCode.Uniqueidentifier)
            {
                return new UniqueIdentifierValidator(attr.LogicalName, attr);
            }
            if (attr.AttributeType.Value == AttributeTypeCode.Picklist)
            {
                return new OptionsetValidator(attr.LogicalName, attr);
            }
            // unfortunately can't validate this....
            //if (attr.AttributeType.Value == AttributeTypeCode.Lookup)
            //{
            //    return new LookupValidator(attr.LogicalName, attr, _crmService);
            //}

            return new NoValidator(attr.LogicalName, attr);
		}
	}
}
