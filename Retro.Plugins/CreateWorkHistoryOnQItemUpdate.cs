using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Retro.Plugins
{
    public class CreateWorkHistoryOnQItemUpdate : IPlugin
    {
        const string Case = "case";
        const string QueueItem = "queueitem";
        public void Execute(IServiceProvider serviceProvider)
        {

            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            ITracingService tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            tracing.Trace(" **************** CreateWorkHistoryOnQItemUpdat *************plugin triggered ");

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && context.PrimaryEntityName.ToUpper().Equals(QueueItem.ToUpper()))
            {
                Entity Target = context.InputParameters["Target"] as Entity;
                Entity PreImage = context.PreEntityImages.Contains("Image") ? (Entity)context.PreEntityImages["Image"] : null;
                Entity PostImage = context.PostEntityImages.Contains("Image") ? (Entity)context.PostEntityImages["Image"] : null;
                string objectType = PreImage.FormattedValues.Contains("objecttypecode") ? PreImage.FormattedValues["objecttypecode"] : string.Empty;
                Entity caseRecord = null;
                Guid caseId;
                try
                {
                    if (string.Equals(Case.ToUpper(), objectType.ToUpper()))
                    {
                        if (PreImage.Attributes.Contains("workerid") && PreImage.GetAttributeValue<EntityReference>("workerid").Id != null && PreImage.GetAttributeValue<EntityReference>("workerid").Id != Guid.Empty)
                        {    // fetch case record
                            caseId = PreImage.Attributes.Contains("objectid") ? PreImage.GetAttributeValue<EntityReference>("objectid").Id : Guid.Empty;
                            caseRecord = service.Retrieve(PreImage.GetAttributeValue<EntityReference>("objectid").LogicalName, caseId, new ColumnSet("new_typeofcase", "statecode"));
                            // check case type == incident

                            // if yes, call createWorkOrder method 

                            //if no return
                            if (caseRecord != null)
                            {
                                if ((caseRecord.FormattedValues.Contains("new_typeofcase") && caseRecord.FormattedValues["new_typeofcase"] == "Incident") && (caseRecord.Attributes.Contains("statecode") && caseRecord.GetAttributeValue<OptionSetValue>("statecode").Value == 0))
                                {
                                    createWorkOrder(service, PreImage, Target, tracing);
                                }
                                else
                                {
                                    tracing.Trace("Case type is not incident / case is not active");
                                    return;
                                }
                            }
                            else
                            {
                                tracing.Trace("Case Record is null");
                                return;
                            }

                        }
                    }
                    else
                    {
                        tracing.Trace("Worked ID is null in preimage");
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

        public void createWorkOrder(IOrganizationService service, Entity PreImage, Entity Target, ITracingService tracing)
        {
            try
            {
                Entity workHistory = new Entity("new_queueworkhistory");
                workHistory["new_timespendbycurrentuser"] = PreImage.Attributes.Contains("new_timespendbycurruser") ? PreImage.GetAttributeValue<decimal>("new_timespendbycurruser") : new decimal(0);
                workHistory["new_timespendbycaseinqueue"] = PreImage.Attributes.Contains("new_timespendoncasequeue") ? PreImage.GetAttributeValue<decimal>("new_timespendoncasequeue") : new decimal(0);
                workHistory["new_totaltimespendoncase"] = PreImage.Attributes.Contains("new_totaltimespendoncase") ? PreImage.GetAttributeValue<decimal>("new_totaltimespendoncase") : new decimal(0);
                if (PreImage.Attributes.Contains("enteredon"))
                {
                    workHistory["new_enteredqueue"] = PreImage.GetAttributeValue<DateTime>("enteredon");
                }
                workHistory["new_queue"] = PreImage.Attributes.Contains("queueid") ? new EntityReference(PreImage.GetAttributeValue<EntityReference>("queueid").LogicalName, PreImage.GetAttributeValue<EntityReference>("queueid").Id) : null;
                workHistory["new_case"] = PreImage.Attributes.Contains("objectid") ? new EntityReference(PreImage.GetAttributeValue<EntityReference>("objectid").LogicalName, PreImage.GetAttributeValue<EntityReference>("objectid").Id) : null;
                workHistory["new_workedby"] = PreImage.Attributes.Contains("workerid") ? new EntityReference("systemuser", PreImage.GetAttributeValue<EntityReference>("workerid").Id) : null;

                Guid NewWorkHistory = service.Create(workHistory);
                tracing.Trace("New work history created with Guid" + NewWorkHistory.ToString());
            }
            catch (Exception ex)
            {

                throw new Exception(ex.ToString());
            }



        }


    }
}
