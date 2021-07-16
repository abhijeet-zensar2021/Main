using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;

namespace HSL.NFUM.PackageDeployer.DataImport.AttributeValidators
{
    
    public class UniqueIdentifierValidator : CrmValidatorBase
    {
        public UniqueIdentifierValidator(string logicalName, AttributeMetadata attr) : base(logicalName, attr)
        {
        }

        public override string[] Validate(string val)
        {
            if (!Guid.TryParse(val, out Guid res))
            {
                return new string[] { $"Error: guid {val} is invalid" };
            }
            
            return new string[] { };
        }
    }
}
