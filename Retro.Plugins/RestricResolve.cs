using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Retro.Plugins
{
    public class RestricResolve : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            ITracingService tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            if (context.InputParameters.Contains("IncidentResolution"))
            {
                Entity incidentResolution = (Entity)context.InputParameters["IncidentResolution"];
                Guid relatedIncidentGuid = ((EntityReference)incidentResolution.Attributes["incidentid"]).Id;

                // Obtain the organization service reference.

                bool HaschildCase = CheckChildCase(relatedIncidentGuid, service);
                if (HaschildCase)
                {
                    throw new InvalidPluginExecutionException(OperationStatus.Succeeded, "It has One or More Child Cases in Active state, hence cannot resolve the parent case");
                }
            }
        }

        private bool CheckChildCase(Guid relatedIncidentGuid, IOrganizationService service)
        {
            string FetchCase = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='incident'>
                                <attribute name='incidentid' />
                                <attribute name='caseorigincode' />
                                <order attribute='title' descending='false' />
                                <filter type='and'>
                                    <condition attribute='parentcaseid' operator='eq' value='{0}' />
                                    <condition attribute='statuscode' operator='eq' value='1' />
                                </filter>
                                </entity>
                            </fetch>";
            EntityCollection ChildColl = service.RetrieveMultiple(new FetchExpression(string.Format(FetchCase, relatedIncidentGuid)));
            if (ChildColl != null && ChildColl.Entities != null)
                return (ChildColl.Entities.Count > 0);
            else
                return false;
        }
    }
}