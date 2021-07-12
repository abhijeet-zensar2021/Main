using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;

namespace CRMPackageDeployment.DataImport.AttributeProcessor
{
    public class LookupProcessor : CrmRowProcessor
    {
        public LookupProcessor(string logicalName, AttributeMetadata attr) : base(logicalName, attr) { }
        public override void AddValue(ref Entity record, string val)
        {
            record[this.LogicalName] = val;
        }
    }
}
