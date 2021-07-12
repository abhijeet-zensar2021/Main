using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using System;

namespace CRMPackageDeployment.DataImport.AttributeProcessor
{
   public class UniqueIdentifierProcessor : CrmRowProcessor
    {
        public UniqueIdentifierProcessor(string logicalName, AttributeMetadata attr) : base(logicalName, attr) { }
        public override void AddValue(ref Entity record, string val)
        {
            throw new NotImplementedException();
        }
    }
}
