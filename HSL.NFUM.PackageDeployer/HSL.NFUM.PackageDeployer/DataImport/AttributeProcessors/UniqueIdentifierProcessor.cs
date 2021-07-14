using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HSL.NFUM.PackageDeployer.DataImport.AttributeProcessors
{
    public class UniqueIdentifierProcessor : CrmRowProcessor
    {
        public UniqueIdentifierProcessor(string logicalName, AttributeMetadata attr) : base(logicalName, attr) { }

        public override void AddValue(ref Entity record, string val)
        {
            record[this.LogicalName] = Guid.Parse(val);
        }
    }
}
