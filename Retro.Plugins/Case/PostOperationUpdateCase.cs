using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Retro.Plugins.Common;
using System;
using System.Linq;
namespace Retro.Plugins
{
    public class PostOperationUpdateCase : IPlugin
    {
        Guid CaseId = Guid.Empty;       
        const string Case = "case";
        const string Active = "active";
        const string ClosedID = "0e78cefb-6f3f-eb11-a813-000d3a18ee0f";
        const string ActiveID = "a0aba2da-6f3f-eb11-a813-000d3a18ee0f";
        const string Closed = "closed";
        const string IncidentCaseType = "incident";

        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            tracing.Trace(" **************** CreateWorkHistoryOnQItemUpdat *************plugin triggered ");

            BusinessLogic CommonLogic = new BusinessLogic();
       
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    Entity CaseTarget = (Entity)context.InputParameters["Target"];
                    //Entity CasePreImage = context.PreEntityImages["PreImage"];
                    Entity Case = CommonLogic.RecordFetch(service, CaseTarget.LogicalName, CaseTarget.Id, Modal.caseColumns);

                    if (Case.FormattedValues["cr32a_typeofcase"].ToString().ToUpper().Equals(IncidentCaseType.ToUpper()))
                    {
                        CaseId = CaseTarget.Id;
                        tracing.Trace("before fetching previous workHistory method");
                        Entity _prevWorkHistory = CommonLogic.FetchPreviousWorkHistory(service, CaseId, tracing);
                        tracing.Trace("after fetching previous workHistory method");
                        tracing.Trace("before calling update workHistory method");
                        //check if status is resolved in context.

                        // Assgin, Update , create
                        //preimage[ownerid]
                        //targ[ownerid]
                        
                        if(CaseTarget.Attributes.Contains("ownerid"))
                        {
                            tracing.Trace("Assign Scenario Entering");
                            bool IsMadeInactive = CommonLogic.UpdateWorkHistory(_prevWorkHistory, service, Case, tracing);
                            tracing.Trace("after calling update workHistory method");
                            if (IsMadeInactive)
                            {
                                tracing.Trace("before calling create workHistory method");
                                Guid NewWorkHistoryID = CommonLogic.CreateWorkHistory(_prevWorkHistory,
                                    Case, service);
                                tracing.Trace("Created new Work history:" + NewWorkHistoryID.ToString());
                                tracing.Trace("after calling create workHistory method");
                            }
                        }
                        else if(CaseTarget.Attributes.Contains("cr32a_casestatusreason"))
                        {
                            string CaseStatusID = CaseTarget.GetAttributeValue<EntityReference>("cr32a_casestatusreason").Id.ToString();
                            if (CaseStatusID.ToUpper().Equals(ClosedID.ToUpper()))
                            {
                                //resolved or closed
                                CommonLogic.UpdateWorkHistory(_prevWorkHistory, service, CaseTarget, tracing);
                            }
                            else if (CaseStatusID.ToUpper().Equals(ActiveID.ToUpper()))
                            {
                                //reopened - create a  new work history
                                CommonLogic.CreateWorkHistory(_prevWorkHistory,
                                    Case, service);                                
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

        

    }
}
