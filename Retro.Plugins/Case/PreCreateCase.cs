using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Retro.Plugins.Common;
using System;
using System.Linq;

namespace Retro.Plugins.Case
{
    public class PreCreateCase : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            tracing.Trace(" **************** PreCreateCase *************plugin triggered ");

            BusinessLogic CommonLogic = new BusinessLogic();
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                try
                {
                    Entity CaseTarget = (Entity)context.InputParameters["Target"];
                    tracing.Trace("inside PreCreate plugin");
                    tracing.Trace("inside Id" + CaseTarget.Id);
                    string orgUrl = GetOrgName(service, tracing, context.OrganizationId);
                    // var orgName = service.ConnectedOrgPublishedEndpoints[Microsoft.Xrm.Sdk.Discovery.EndpointType.WebApplication];
                    string entityLogicalName = context.PrimaryEntityName;
                    string Id = Convert.ToString(CaseTarget.Id);
                    string appId = GetAppId(service, tracing);

                    string url = string.Concat("https://", orgUrl, ".crm.dynamics.com/main.aspx?" +
                        "appid=", appId,
                        "&pagetype=entityrecord",
                        "&etn=", context.PrimaryEntityName,
                        "&id=", CaseTarget.Id);

                    CaseTarget["description"] = url;
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(string.Concat(ex.Message, "   stack trace:", ex.StackTrace), ex);
                }

                //@"https://myorg.crm.dynamics.com/main.aspx?etn=account&pagetype=entityrecord&id=%7B91330924-802A-4B0D-A900-34FD9D790829%7D";
            }
        }

        private string GetAppId(IOrganizationService service, ITracingService tracing)
        {
            // Define Condition Values
            string QEappmodule_name = "Ms_learn";

            // Instantiate QueryExpression QEappmodule
            QueryExpression QEappmodule = new QueryExpression("appmodule");

            // Add columns to QEappmodule.ColumnSet
            QEappmodule.ColumnSet.AddColumns("publisherid", "appmoduleidunique", "name", "organizationid", "appmoduleid");

            // Define filter QEappmodule.Criteria
            QEappmodule.Criteria.AddCondition("name", ConditionOperator.Equal, QEappmodule_name);

            EntityCollection Collection = service.RetrieveMultiple(QEappmodule);
            return Collection.Entities.FirstOrDefault().Attributes.Contains("appmoduleidunique") ? Convert.ToString(Collection.Entities.FirstOrDefault().Id) : "";
        }

        private string GetOrgName(IOrganizationService service, ITracingService tracing, Guid organizationId)
        {
            // Define Condition Values
            Entity Org = service.Retrieve("organization", organizationId, new ColumnSet("name"));

            return Org["name"].ToString();

            QueryExpression organizationquery = new QueryExpression("organization")
            {
                ColumnSet = new ColumnSet("name")
            };
            Entity organization = service.RetrieveMultiple(organizationquery).Entities.FirstOrDefault();
            string organizationname = organization.GetAttributeValue<string>("name");
            EntityCollection Collection = service.RetrieveMultiple(organizationquery);
            return Collection.Entities.FirstOrDefault().Attributes.Contains("name") ? Collection.Entities.FirstOrDefault()["name"].ToString() : "";
        }
    }
}