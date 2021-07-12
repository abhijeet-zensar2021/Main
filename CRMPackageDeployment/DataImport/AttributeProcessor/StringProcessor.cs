using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;

namespace CRMPackageDeployment.DataImport.AttributeProcessor
{
    public class StringProcessor : CrmRowProcessor
    {
        public StringProcessor (string logicalName, AttributeMetadata attr): base(logicalName, attr) { }
        public override void AddValue(ref Entity record, string val)
        {
            record[this.LogicalName] = val;
        }
    }
}
