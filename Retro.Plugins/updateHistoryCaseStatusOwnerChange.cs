using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
namespace Retro.Plugins
{
    public class updateHistoryCaseStatusOwnerChange : IPlugin
    {
        private Guid CaseId = Guid.Empty;
        const string Case = "case";
        const string Active = "active";
        const string ClosedID = "95d97865-e93d-eb11-a813-000d3ac9ccda";
        const string ActiveID = "b46dd35d-e93d-eb11-a813-000d3ac9ccda";
        const string Closed = "closed";
        const string IncidentCaseType = "incident";

        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            tracing.Trace(" **************** CreateWorkHistoryOnQItemUpdat *************plugin triggered ");

       
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    Entity CaseTarget = (Entity)context.InputParameters["Target"];
                    Entity CasePreImage = context.PreEntityImages["PreImage"];

                    if (CasePreImage.FormattedValues["new_typeofcase"].ToString().ToUpper().Equals(IncidentCaseType.ToUpper()))
                    {
                        CaseId = CaseTarget.Id;
                        tracing.Trace("before fetching previous workHistory method");
                        Entity _prevWorkHistory = FetchPreviousWorkHistory(service, CaseId, tracing);
                        tracing.Trace("after fetching previous workHistory method");
                        tracing.Trace("before calling update workHistory method");
                        //check if status is resolved in context.

                        // Assgin, Update , create
                        //preimage[ownerid]
                        //targ[ownerid]

                        if(CaseTarget.Attributes.Contains("ownerid")&& CaseTarget.GetAttributeValue<EntityReference>("ownerid").Id != CasePreImage.GetAttributeValue<EntityReference>("ownerid").Id)
                        {
                            tracing.Trace("Assign Scenario Entering");
                            bool IsMadeInactive = UpdateWorkHistory(_prevWorkHistory, service, FetchCaseStatus(CaseTarget, CasePreImage), tracing);
                            tracing.Trace("after calling update workHistory method");
                            if (IsMadeInactive)
                            {
                                tracing.Trace("before calling create workHistory method");
                                Guid NewWorkHistoryID = CreateWorkHistory(_prevWorkHistory, FetchCaseOwner(CaseTarget, CasePreImage), FetchCaseStatus(CaseTarget, CasePreImage), service);
                                tracing.Trace("Created new Work history:" + NewWorkHistoryID.ToString());
                                tracing.Trace("after calling create workHistory method");
                            }
                        }
                        else if(CaseTarget.Attributes.Contains("dev_casetypestatus"))
                        {
                            string CaseStatusID = CaseTarget.GetAttributeValue<EntityReference>("dev_casetypestatus").Id.ToString();
                            if (CaseStatusID.ToUpper().Equals(ClosedID.ToUpper()))
                            {
                                //resolved or closed
                                UpdateWorkHistory(_prevWorkHistory, service, FetchCaseStatus(CaseTarget, CasePreImage), tracing);
                            }
                            else if (CaseStatusID.ToUpper().Equals(ActiveID.ToUpper()))
                            {
                                //reopened - create a  new work history
                                CreateWorkHistory(_prevWorkHistory, FetchCaseOwner(CaseTarget, CasePreImage), FetchCaseStatus(CaseTarget, CasePreImage), service);                                
                            }

                        }



                        #region Commented 

                        /*
                                                if (CaseTarget.Attributes.Contains("dev_casetypestatus"))
                                                {
                                                    string CaseStatus = CaseTarget.GetAttributeValue<EntityReference>("dev_casetypestatus").Name;
                                                    if (CaseStatus.ToUpper().Equals(Closed.ToUpper()))
                                                    {
                                                        //resolved 
                                                        UpdateWorkHistory(_prevWorkHistory, service, FetchCaseStatus(CaseTarget, CasePreImage), tracing);
                                                    }
                                                    else if (CaseStatus.ToUpper().Equals(Active.ToUpper()))
                                                    {
                                                        CreateWorkHistory(_prevWorkHistory, FetchCaseOwner(CaseTarget, CasePreImage), FetchCaseStatus(CaseTarget, CasePreImage), service);
                                                        //reopened - create a  new work history
                                                    }
                                                }
                                                else if (CaseTarget.Attributes.Contains("ownerid") || (CaseTarget.Attributes.Contains("ownerid") && ((CaseTarget.Attributes.Contains("dev_casetypestatus") &&
                                                    (CaseTarget.GetAttributeValue<EntityReference>("dev_casetypestatus").Name.ToUpper().Equals(Active.ToUpper()))))))
                                                {
                                                    bool IsMadeInactive = UpdateWorkHistory(_prevWorkHistory, service, FetchCaseStatus(CaseTarget, CasePreImage), tracing);
                                                    tracing.Trace("after calling update workHistory method");
                                                    if (IsMadeInactive)
                                                    {
                                                        tracing.Trace("before calling create workHistory method");
                                                        Guid NewWorkHistoryID = CreateWorkHistory(_prevWorkHistory, FetchCaseOwner(CaseTarget, CasePreImage), FetchCaseStatus(CaseTarget, CasePreImage), service);
                                                        tracing.Trace("Created new Work history:" + NewWorkHistoryID.ToString());
                                                        tracing.Trace("after calling create workHistory method");
                                                    }
                                                }*/

                        #endregion
                    }
                    else
                    {
                        return;
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

        private Guid CreateWorkHistory(Entity _preWorkHistory, EntityReference CaseOwner, EntityReference CaseStatus, IOrganizationService service)
        {
            Entity _newWorkHistory = new Entity(_preWorkHistory.LogicalName);

            if (CaseOwner != null)
            {
                if (CaseOwner.LogicalName.ToLower().Equals("systemuser"))
                {
                    _newWorkHistory["new_workedby"] = new EntityReference(CaseOwner.LogicalName, CaseOwner.Id);
                }
                else
                {
                    _newWorkHistory["dev_workedbyteam"] = new EntityReference(CaseOwner.LogicalName, CaseOwner.Id);
                }
            }
            if (CaseStatus != null)
            {
                _newWorkHistory["dev_casestatusreason"] = new EntityReference(CaseStatus.LogicalName, CaseStatus.Id);
            }// _newWorkHistory["statuscode"] = new OptionSetValue(1);//1 - active

            _newWorkHistory["new_case"] = _preWorkHistory["new_case"];
            _newWorkHistory["new_name"] = _preWorkHistory["new_name"];
            Guid _newID = service.Create(_newWorkHistory);
            return _newID;
        }
        internal bool UpdateWorkHistory(Entity _preWorkHistory, IOrganizationService service, EntityReference CaseStatus, ITracingService tracing)
        {
            tracing.Trace("inside UpdateWorkHistory method");
            bool isUpdated = false;
            //inactivate the work history record
            _preWorkHistory["new_timespendbycurrentuser"] = _preWorkHistory.Attributes.Contains("new_calctimespendbycurruser") ? _preWorkHistory.GetAttributeValue<decimal>("new_calctimespendbycurruser") : new decimal(0.00);
            _preWorkHistory["new_timespendbycaseinqueue"] = _preWorkHistory.Attributes.Contains("new_calctimespendqueue") ? _preWorkHistory.GetAttributeValue<decimal>("new_calctimespendqueue") : new decimal(0.00);
            _preWorkHistory["new_totaltimespendoncase"] = _preWorkHistory.Attributes.Contains("new_calctotaltimespendbycurruser") ? _preWorkHistory.GetAttributeValue<decimal>("new_calctotaltimespendbycurruser") : new decimal(0.00);
            if (CaseStatus != null)
            {
                _preWorkHistory["dev_casestatusreason"] = new EntityReference(CaseStatus.LogicalName, CaseStatus.Id);
            }
            _preWorkHistory["statecode"] = new OptionSetValue(1);//2 - Inactive
            _preWorkHistory["statuscode"] = new OptionSetValue(2);//2 - Inactive
            service.Update(_preWorkHistory);
            isUpdated = true;
            tracing.Trace("Updated existing  Work history:" + _preWorkHistory.Id.ToString());
            return isUpdated;
        }
        private Entity FetchPreviousWorkHistory(IOrganizationService service, Guid CaseId, ITracingService tracing)
        {
            string FetchWorkHistory = $@"<fetch >
                                          <entity name='new_queueworkhistory' >
                                            <attribute name='new_workedby' />
                                            <attribute name='new_case' />
                                            <attribute name='new_name' />
                                            <attribute name='new_calctimespendbycurruser' />
                                            <attribute name='new_calctotaltimespendbycurruser' />
                                            <attribute name='new_casestatus' />
                                            <attribute name='dev_workedbyteam' />
                                            <attribute name='new_calctimespendqueue' />
                                            <filter type='and' >
                                              <condition attribute='new_case' operator='eq' value='{CaseId}' />
                                              <condition attribute='statuscode' operator='eq' value='1' />
                                              <condition attribute='statecode' operator='eq' value='0' />
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

        private EntityReference FetchCaseStatus(Entity Target, Entity PreImage)
        {
            EntityReference CaseStatus = null;

            if (Target.Attributes.Contains("dev_casetypestatus"))
            {
                CaseStatus = new EntityReference(Target.GetAttributeValue<EntityReference>("dev_casetypestatus").LogicalName, Target.GetAttributeValue<EntityReference>("dev_casetypestatus").Id);
            }
            else
            {
                CaseStatus = new EntityReference(PreImage.GetAttributeValue<EntityReference>("dev_casetypestatus").LogicalName, PreImage.GetAttributeValue<EntityReference>("dev_casetypestatus").Id);
            }
            return CaseStatus;

        }

        private EntityReference FetchCaseOwner(Entity Target, Entity PreImage)
        {
            EntityReference CaseOwner = null;

            if (Target.Attributes.Contains("ownerid"))
            {
                CaseOwner = new EntityReference(Target.GetAttributeValue<EntityReference>("ownerid").LogicalName, Target.GetAttributeValue<EntityReference>("ownerid").Id);
            }
            else
            {
                CaseOwner = new EntityReference(PreImage.GetAttributeValue<EntityReference>("ownerid").LogicalName, PreImage.GetAttributeValue<EntityReference>("ownerid").Id);
            }
            return CaseOwner;

        }

    }
}
