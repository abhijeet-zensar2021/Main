using HSL.NFUM.PackageDeployer.DataImport.AttributeProcessors;
using HSL.NFUM.PackageDeployer.DataImport.Extensions;
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

    public class CrmCalendarImport
	{
		private readonly IOrganizationService _crmService;
		private readonly EntityMetadata _metadata;
		private readonly SortedDictionary<Guid, Entity> _allRecords;
		private readonly CrmRowProcessor[] _rowProcessors;
        private readonly ITraceLogging _log;

        public CrmCalendarImport(IOrganizationService crmService, EntityMetadata metadata, CrmRowProcessor[] rowProcessors, ITraceLogging log)
		{
			_crmService = crmService;
			_metadata = metadata;
			_rowProcessors = rowProcessors;
            _log = log;

            // retrieve the empty calendars, these will be used for the updates
            var allEntityQuery = new QueryExpression("calendar")
			{
				ColumnSet = new ColumnSet("name")
			};
			var records = crmService.RetrieveAllRecords(allEntityQuery).ToDictionary(x => x.Id);
			_allRecords = new SortedDictionary<Guid, Entity>(records);
		}

		internal void UpsertRows(List<string[]> rows)
		{
			// create and populate the required entities
			var requiresUpdates = new HashSet<Guid>();
			foreach (var row in rows)
			{
				var ent = new Entity(_metadata.LogicalName);
				for (var i = 0; i < _rowProcessors.Count(); i++)
				{
					_rowProcessors[i].AddValue(ref ent, row[i]);
					if (_rowProcessors[i].LogicalName.Equals(_metadata.PrimaryIdAttribute, StringComparison.OrdinalIgnoreCase))
					{
						ent.Id = (Guid)ent[_rowProcessors[i].LogicalName];
					}
				}

				AddDefaults(ent);

				// get the corrosponsing calendar
				var calId = ent.GetAttributeValue<EntityReference>("calendarid");
				var matchingCalendar = _allRecords[calId.Id];

				var calRules = matchingCalendar.GetAttributeValue<EntityCollection>("calendarrules").Entities;
				var matchingRule = calRules.FirstOrDefault(x => x.Id.Equals(ent.GetAttributeValue<Guid>("calendarruleid")));

				if (matchingRule == null)
				{
					// add the rule and signal the update to occur to the calendar item
					_log.Log($"Rule created: {ent.Id}", TraceEventType.Information);
					calRules.Add(ent);
					requiresUpdates.Add(calId.Id);
				}
				else
				{
					// replace the attributes if the entities require updating
					if (RuleRequiresUpdate(ent, matchingRule))
					{
						matchingRule.Attributes = ent.Attributes;
						requiresUpdates.Add(calId.Id);
					}
				}
			}

			// process the updates
			foreach (var calId in requiresUpdates)
            {
				_log.Log($"Updating {calId} : {_allRecords[calId].GetAttributeValue<string>("name")}", TraceEventType.Information);
				_crmService.Update(_allRecords[calId]);
            }
		}

		/// <summary>
		/// This is a special case so we'll include the defaults that we need
		/// </summary>
		/// <param name="ent"></param>
        private void AddDefaults(Entity ent)
        {
			// else it seems to import as UTC (which i didn't think CRM was supposed to do?)
			var date = (DateTime)ent["starttime"];
			if (date.IsDaylightSavingTime())
			{
				ent["starttime"] = date.AddHours(1);
			}

			ent["pattern"] = "FREQ=DAILY;INTERVAL=1;COUNT=1";
			ent["issimple"] = false;
			ent["rank"] = 0;
			ent["duration"] = 1440;
			ent["timezonecode"] = -1;
			ent["timecode"] = 2;
			ent["subcode"] = 5;
			ent["effectiveintervalstart"] = ent["starttime"];
			ent["effectiveintervalend"] = ent.GetAttributeValue<DateTime>("starttime").AddDays(1);
		}

        private bool RuleRequiresUpdate(Entity ent, Entity currentRule)
		{
			// compare to work out if requires updating
			var currentRecord = currentRule;
			var differingAttributes = _rowProcessors
                .Where(rp => 
					(!(ent[rp.LogicalName] is DateTime) && !ent[rp.LogicalName].Equals(currentRecord[rp.LogicalName])) ||

					// when comparing datetime source -> to universal, current -> to local 
					((ent[rp.LogicalName] is DateTime) && !ent.GetAttributeValue<DateTime>(rp.LogicalName).Equals(currentRecord.GetAttributeValue<DateTime>(rp.LogicalName).ToLocalTime())) 
					);

			if (differingAttributes.Any())
			{
				var attr = differingAttributes.First().LogicalName;
				_log.Log($"Rule updated: {currentRule.Id} {ent[attr]} {currentRule[attr]} ", TraceEventType.Information);
				return true;
			}

			_log.Log($"Rule no change: {currentRule.Id}", TraceEventType.Information);
			return false;
		}
	}
}
