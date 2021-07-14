using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using System;

namespace HSL.NFUM.PackageDeployer.DataImport.AttributeValidators
{
    public class NoValidator : CrmValidatorBase
    {
        public NoValidator(string logicalName, AttributeMetadata attr) : base(logicalName, attr)
        {
        }

        public override string[] Validate(string val)
        {
            return new string[] { };
        }
    }
}
