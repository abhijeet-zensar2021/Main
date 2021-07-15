using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;

namespace HSL.NFUM.PackageDeployer.DataImport.AttributeProcessors
{
    public class StringProcessor : CrmRowProcessor
    {
        public StringProcessor(string logicalName, AttributeMetadata attr) : base(logicalName, attr) { }

        public override void AddValue(ref Entity record, string val)
        {
            record[this.LogicalName] = val;
        }
    }
}
