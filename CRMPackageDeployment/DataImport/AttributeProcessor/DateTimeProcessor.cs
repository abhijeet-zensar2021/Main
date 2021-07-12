using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using System;

namespace CRMPackageDeployment.DataImport.AttributeProcessor
{
    public class DateTimeProcessor : CrmRowProcessor
    {
        public DateTimeProcessor(string logicalName, AttributeMetadata attr) : base(logicalName, attr) { }
        public override void AddValue(ref Entity record, string val)
        {
            record[this.LogicalName] = DateTime.Parse(val);
        }
    }
}
