using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using System;

namespace CRMPackageDeployment.DataImport.AttributeProcessor
{
    public class BooleanProcessor : CrmRowProcessor
    {
        private readonly BooleanAttributeMetadata _booleanMetadata;
        public BooleanProcessor(string logicalName, AttributeMetadata attr) : base(logicalName, attr)
        {
            _booleanMetadata = (BooleanAttributeMetadata)attr;
        }
        public override void AddValue(ref Entity record, string val)
        {
            var isTrue = _booleanMetadata.OptionSet.TrueOption.Label.UserLocalizedLabel.Label.Equals(val);
            var isFalse = _booleanMetadata.OptionSet.FalseOption.Label.UserLocalizedLabel.Label.Equals(val);
            if(!isTrue && !isFalse)
            {
                throw new Exception("Error: Could not find boolean value");
            }
            record[this.LogicalName] = isTrue;
        }
    }
}
