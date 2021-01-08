using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Retro.Plugins.Common;
using System;
using System.Linq;

namespace Retro.Plugins
{
    public class CreateWorkHistoryOnQItemUpdate : IPlugin
    {
        private const string Case = "case";
        private const string QueueItem = "queueitem";

        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            BusinessLogic common = new BusinessLogic(serviceFactory,tracing);
            tracing.Trace(" **************** CreateWorkHistoryOnQItemCreate *************plugin triggered ");

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && context.PrimaryEntityName.ToUpper().Equals(QueueItem.ToUpper()))
            {
                Entity queueItem = context.InputParameters["Target"] as Entity;
                string objectType = queueItem.FormattedValues.Contains("objecttypecode") ? queueItem.FormattedValues["objecttypecode"] : string.Empty;
                Entity caseRecord = null;
                string queueName = queueItem.Attributes.Contains("queueid") ? common.FetchQueueName(((EntityReference)queueItem["queueid"]).Id) : "";
                tracing.Trace("queue name {0}", queueName);
                caseRecord = FetchCase(queueItem, service);
                try
                {
                    if (string.Equals(Case.ToUpper(), objectType.ToUpper()) && caseRecord != null)
                    {
                        //fetch previous workhistory record
                        tracing.Trace("Before fetch previous record with active status");
                        Entity preWorkHistory = common.FetchPreviousWorkHistory(service, caseRecord.Id, tracing);
                        tracing.Trace("Before updating previous record");
                        // update previous workhistory record
                       // bool updateWH = common.UpdateWorkHistory(preWorkHistory, caseRecord, tracing);
                        //if (updateWH && !string.IsNullOrEmpty(queueName))
                        //{
                            tracing.Trace("Before creating new record with active status");
                            // common.CreateWorkHistory(preWorkHistory, caseRecord, service, QueueTarget);
                            common.updateCaseOwner(caseRecord, queueName);
                        //}
                        //create new one
                    }
                    else
                    {
                        tracing.Trace("Object type is not case ");
                        return;
                    }
                }
                catch (Exception ex)
                {
                    tracing.Trace("error has worked " + ex.ToString());
                    //  throw;
                }
            }
            else
            {
                tracing.Trace("Target is not an Entity ");
                return;
            }
        }

        private Entity FetchCase(Entity target, IOrganizationService service)
        {
            Entity caseRecord = null;
            ConditionExpression condition1 = new ConditionExpression();
            condition1.AttributeName = "incidentid";
            condition1.Operator = ConditionOperator.Equal;
            condition1.Values.Add(target.GetAttributeValue<EntityReference>("objectid").Id);
            FilterExpression filter1 = new FilterExpression();
            filter1.Conditions.Add(condition1);
            QueryExpression query = new QueryExpression("incident");
            query.ColumnSet.AddColumns(Common.Modal.caseColumns);
            query.Criteria.AddFilter(filter1);
            EntityCollection incidentCollection = service.RetrieveMultiple(query);
            if (incidentCollection.Entities.Count > 0)
            {
                caseRecord = incidentCollection.Entities.FirstOrDefault();
            }
            return caseRecord;
        }
    }
}