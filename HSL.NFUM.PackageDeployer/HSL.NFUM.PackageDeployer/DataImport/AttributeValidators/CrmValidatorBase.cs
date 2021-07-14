using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;

namespace HSL.NFUM.PackageDeployer.DataImport.AttributeValidators
{
    public abstract class CrmValidatorBase
    {
        public readonly string LogicalName;
        protected readonly AttributeMetadata Attribute;

        public CrmValidatorBase(string logicalName, AttributeMetadata attr)
        {
            LogicalName = logicalName;
            Attribute = attr;
        }

        // Mutates the record, add the item
        public abstract string[] Validate(string val);

        // can be overridden for the custom types
        public virtual string[] GetColumnsetColumns()
        {
            return new[] { LogicalName };
        }
    }
}
