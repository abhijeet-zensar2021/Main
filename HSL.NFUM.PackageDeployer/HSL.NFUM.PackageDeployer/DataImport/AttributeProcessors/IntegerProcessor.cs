using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using System;

namespace HSL.NFUM.PackageDeployer.DataImport.AttributeProcessors
{
    public class IntegerProcessor : CrmRowProcessor
    {

        public IntegerProcessor(string logicalName, AttributeMetadata attr) : base(logicalName, attr)
        {
        }

        public override void AddValue(ref Entity record, string val)
        {
            record[this.LogicalName] = int.Parse(val);
        }
    }
}
