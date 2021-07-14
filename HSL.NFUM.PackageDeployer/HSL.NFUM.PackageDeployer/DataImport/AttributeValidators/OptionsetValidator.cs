using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Linq;

namespace HSL.NFUM.PackageDeployer.DataImport.AttributeValidators
{
    public class OptionsetValidator : CrmValidatorBase
    {
        private PicklistAttributeMetadata _picklistMetadata;

        public OptionsetValidator(string logicalName, AttributeMetadata attr) : base(logicalName, attr)
        {
            _picklistMetadata = (PicklistAttributeMetadata)attr;
        }

        public override string[] Validate(string val)
        {
            var optionset = _picklistMetadata.OptionSet.Options.SingleOrDefault(o => o.Label.LocalizedLabels.Any(l => l.Label.Equals(val, StringComparison.OrdinalIgnoreCase)));
            if (optionset == null)
            {
                return new string[] { $"Error: No matching option for label {val} on {Attribute.EntityLogicalName}.{Attribute.LogicalName}" };
            }

            return new string[] { };
        }
    }
}
