using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;

namespace HSL.NFUM.PackageDeployer.DataImport.AttributeValidators
{
    public class StringValidator : CrmValidatorBase
    {
        public StringValidator(string logicalName, AttributeMetadata attr) : base(logicalName, attr)
        {
        }

        public override string[] Validate(string val)
        {
            var attribute = (StringAttributeMetadata)this.Attribute;
            if (!string.IsNullOrEmpty(val) && val.Length > attribute.MaxLength)
            {
                return new string[] { $"{val} is greater then max length for {this.Attribute.EntityLogicalName}.{this.LogicalName}" };
            }

            return new string[] { };
        }
    }
}
