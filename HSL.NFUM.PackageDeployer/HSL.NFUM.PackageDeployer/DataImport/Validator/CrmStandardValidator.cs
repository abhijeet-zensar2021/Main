using HSL.NFUM.PackageDeployer.DataImport.AttributeProcessors;
using HSL.NFUM.PackageDeployer.DataImport.AttributeValidators;
using HSL.NFUM.PackageDeployer.DataImport.Interfaces;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace HSL.NFUM.PackageDeployer.DataImport.Import
{
    public class CrmStandardValidator
    {
        private readonly IOrganizationService _crmService;
        private readonly EntityMetadata _metadata;
        private readonly CrmValidatorBase[] _rowProcessors;
        private readonly ITraceLogging _log;

        public CrmStandardValidator(IOrganizationService crmService, EntityMetadata metadata, CrmValidatorBase[] rowProcessors, ITraceLogging log)
        {
            _crmService = crmService;
            _metadata = metadata;
            _rowProcessors = rowProcessors;
            _log = log;
        }

        internal string[] ValidateRow(string[] row)
        {
            var errors = new List<string>();

            // create and populate the rquired entity
            for (var i = 0; i < _rowProcessors.Count(); i++)
            {
                var message = _rowProcessors[i].Validate(row[i]);
                if (message.Length != 0)
                {
                    errors.AddRange(message);
                }
            }

            return errors.ToArray();
        }
    }
}
