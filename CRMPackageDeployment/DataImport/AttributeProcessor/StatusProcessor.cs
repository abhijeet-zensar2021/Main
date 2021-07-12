using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;

namespace CRMPackageDeployment.DataImport.AttributeProcessor
{
    class StatusProcessor : CrmRowProcessor
    {
        public StatusProcessor(string logicalName, AttributeMetadata attr) : base(logicalName, attr) { }
        public override void AddValue(ref Entity record, string val)
        {
            record[this.LogicalName] = val;
        }
    }
}
