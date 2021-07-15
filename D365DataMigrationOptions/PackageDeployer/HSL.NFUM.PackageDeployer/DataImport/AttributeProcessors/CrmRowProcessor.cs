using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;

namespace HSL.NFUM.PackageDeployer.DataImport.AttributeProcessors
{
    public abstract class CrmRowProcessor
    {
        public readonly string LogicalName;
        protected readonly AttributeMetadata Attribute;

        public CrmRowProcessor(string logicalName, AttributeMetadata attr)
        {
            LogicalName = logicalName;
            Attribute = attr;
        }

        // Mutates the record, add the item
        public abstract void AddValue(ref Entity record, string val);

        // can be overridden for the custom types
        public virtual string[] GetColumnsetColumns()
        {
            return new[] { LogicalName };
        }
    }
}
