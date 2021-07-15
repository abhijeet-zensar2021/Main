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
    public class StatusProcessor : CrmRowProcessor
    {
        private readonly StatusAttributeMetadata _statusMetadata;

        public StatusProcessor(string logicalName, AttributeMetadata attr) : base(logicalName, attr)
        {
            _statusMetadata = (StatusAttributeMetadata)attr;
        }

        public override string[] GetColumnsetColumns()
        {
            return new[] {
            this.LogicalName,
            "statecode"
        };
        }

        public override void AddValue(ref Entity record, string val)
        {
            // ignore if null
            if (String.IsNullOrEmpty(val))
            {
                throw new Exception("statuscode attributes requires [" + String.Join(",", _statusMetadata.OptionSet.Options.Select(a => a.Label.LocalizedLabels.FirstOrDefault()?.Label)) + "]");
            }

            var optionset = _statusMetadata.OptionSet.Options.SingleOrDefault(o => o.Label.LocalizedLabels.Any(l => l.Label.Equals(val, StringComparison.OrdinalIgnoreCase)));
            if (optionset == null)
            {
                throw new Exception("Error: could not find matching string");
            }

            record[this.LogicalName] = new OptionSetValue(optionset.Value.Value);

            // add the statecode in
            var statecode = ((StatusOptionMetadata)optionset).State;
            if (statecode.HasValue)
            {
                record["statecode"] = new OptionSetValue(statecode.Value);
            }
        }
    }
}
