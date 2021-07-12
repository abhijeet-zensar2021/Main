using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRMPackageDeployment.DataImport.AttributeProcessor
{
    public abstract class CrmRowProcessor
    {
        public readonly string LogicalName;
        protected readonly AttributeMetadata Attribute;

        public CrmRowProcessor(string logicalName, AttributeMetadata attr)
        {
            LogicalName = logicalName;
            Attribute = attr;
        }

        public abstract void AddValue(ref Entity record, string val);
        public virtual string[] GetColumnsetColumns()
        {
            return new[] { LogicalName };
        }
    }
}
