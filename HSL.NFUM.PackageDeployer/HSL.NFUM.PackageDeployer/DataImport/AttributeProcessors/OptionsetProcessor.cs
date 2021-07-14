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
    public class OptionsetProcessor : CrmRowProcessor
    {
        private readonly PicklistAttributeMetadata _picklistMetadata;

        public OptionsetProcessor(string logicalName, AttributeMetadata attr) : base(logicalName, attr)
        {
            _picklistMetadata = (PicklistAttributeMetadata)attr;
        }

        public override void AddValue(ref Entity record, string val)
        {
            // get the matching label
            var optionset = _picklistMetadata.OptionSet.Options.SingleOrDefault(o => o.Label.LocalizedLabels.Any(l => l.Label.Equals(val, StringComparison.OrdinalIgnoreCase)));
            if (optionset == null)
            {
                throw new Exception("Error: could not find matching string");
            }

            record[this.LogicalName] = new OptionSetValue(optionset.Value.Value);
        }
    }
}
