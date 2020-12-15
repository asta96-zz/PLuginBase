using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Retro.Plugins
{
    public class UpdateWorkHistoryOnQItemCreate : IPlugin
    {
        //const string Case = "case";
        const string QueueItem = "queueitem";

        private Guid CaseId = Guid.Empty;
       

        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            tracing.Trace(" **************** CreateWorkHistoryOnQItemUpdat *************plugin triggered ");

            /*Assign (update of case owner):- 
            //update operation
            Step 1: fetch the previous work history record and all the calculate field values. Filter (status = active) 

            Step 2: Copy the calculated field values and update in the dummy field value in same record. Also update status = inactive . Do Service.Update()

            //Create of new work history record
            Step 3: Create a new work history record, populate owner value = new owner from Target. Do Service.Create()
            */
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && context.PrimaryEntityName.ToUpper().Equals(QueueItem.ToUpper()))
                {
                    Entity Target = context.InputParameters["Target"] as Entity;
                    CaseId = Target.Attributes.Contains("new_case") ? Target.GetAttributeValue<EntityReference>("new_case").Id : Guid.Empty;
                    tracing.Trace("before fetching previous workHistory method");
                    Entity _preWorkHistory = FetchPreviousWorkHistory(service, CaseId, tracing);
                    tracing.Trace("after fetching previous workHistory method");
                    tracing.Trace("before calling update workHistory method");
                    bool IsMadeInactive = UpdateWorkHistory(_preWorkHistory, service, Target, tracing);
                    tracing.Trace("after calling update workHistory method");
                    if (IsMadeInactive)
                    {
                        tracing.Trace("before calling create workHistory method");
                        Guid NewWorkHistoryID = CreateWorkHistory(_preWorkHistory, service);
                        tracing.Trace("Created new Work history:" + NewWorkHistoryID.ToString());
                        tracing.Trace("after calling create workHistory method");
                    }
                }
            }
            catch (Exception ex)
            {
                tracing.Trace("Exception Occured" + ex.Message);
                tracing.Trace("Exception StackTrace" + ex.StackTrace);
                tracing.Trace("InnerException.Message" + ex.InnerException.Message);
                tracing.Trace("InnerException.StackTrace" + ex.InnerException.StackTrace);
            }
        }

        private Guid CreateWorkHistory(Entity _preWorkHistory, IOrganizationService service)
        {
            Entity _newWorkHistory = new Entity(_preWorkHistory.LogicalName);
            _newWorkHistory["statuscode"] = new OptionSetValue(1);//1 - active
            Guid _newID = service.Create(_newWorkHistory);
            return _newID;
        }
        internal bool UpdateWorkHistory(Entity _preWorkHistory, IOrganizationService service, Entity Case, ITracingService tracing)
        {
            tracing.Trace("inside UpdateWorkHistory method");
            bool isUpdated = false;
            //inactivate the work history record
            _preWorkHistory["statecode"] = new OptionSetValue(1);//2 - Inactive
            _preWorkHistory["statuscode"] = new OptionSetValue(2);//2 - Inactive
            service.Update(_preWorkHistory);
            isUpdated = true;
            tracing.Trace("Updated existing  Work history:" + _preWorkHistory.Id.ToString());
            return isUpdated;
         }
        private Entity FetchPreviousWorkHistory(IOrganizationService service,Guid CaseId, ITracingService tracing)
        {
            string FetchWorkHistory = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='new_queueworkhistory'>
                                        <attribute name='new_queueworkhistoryid' />
                                        <attribute name='new_name' />
                                        <attribute name='createdon' />
                                        <order attribute='new_name' descending='false' />
                                        <filter type='and'>
                                          <condition attribute='new_case' operator='eq'value='{{CaseId}}' />
                                          <condition attribute='statecode' operator='eq' value='0' />
                                          <condition attribute='statuscode' operator='eq' value='1' />
                                        </filter>
                                      </entity>
                                    </fetch>";
            Entity History = null;
            EntityCollection coll = service.RetrieveMultiple(new FetchExpression(FetchWorkHistory));
            
            if (coll.Entities.Count > 0)
            {
                tracing.Trace("Fetched WorkHistory Count:" + coll.Entities.Count);
                History = coll.Entities.FirstOrDefault();
            }

            return History;


        }
    }
}
