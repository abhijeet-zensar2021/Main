using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;

namespace CRMPackageDeployment.DataImport.AttributeProcessor
{
    public class OptionsetProcessor : CrmRowProcessor
    {
        public OptionsetProcessor(string logicalName, AttributeMetadata attr) : base(logicalName, attr) { }
        public override void AddValue(ref Entity record, string val)
        {
            record[this.LogicalName] = val;
        }
    }
}
