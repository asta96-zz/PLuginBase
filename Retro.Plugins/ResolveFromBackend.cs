using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System;

namespace Retro.Plugins
{
    public class ResolveFromBackend : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            ITracingService tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            if (context.InputParameters.Contains("caseRecordId"))
            {
                tracing.Trace("contains CaseRecord ID" + ((EntityReference)context.InputParameters["caseRecordId"]).Id);
                tracing.Trace("contains CaseRecord Name" + ((EntityReference)context.InputParameters["caseRecordId"]).Name);

                EntityReference Incident = context.InputParameters["Target"] as EntityReference;
                Entity Incidentresolution = new Entity("incidentresolution");
                Incidentresolution["subject"] = "Closed via backend";
                Incidentresolution["incidentid"] = Incident;

                var closeIncidentRequest = new CloseIncidentRequest
                {
                    IncidentResolution = Incidentresolution,
                    Status = new OptionSetValue(5)
                };

                var closeResponse =
                    (CloseIncidentResponse)service.Execute(closeIncidentRequest);
            }
            else
            {
                tracing.Trace("Doesnot contain caseRecordId");
            }
        }
    }
}