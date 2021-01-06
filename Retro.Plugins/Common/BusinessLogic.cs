using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;

namespace Retro.Plugins.Common
{
    public class BusinessLogic
    {
        public Guid CreateWorkHistory(Entity _preWorkHistory, Entity Case, IOrganizationService service, Entity QueueItem =null)
        {
            Entity _newWorkHistory = new Entity("new_queueworkhistory");
           
            EntityReference CaseStatus = FetchCaseStatus(Case);
            EntityReference Owner = FetchOwner(Case);
            if (CaseStatus != null)
            {
                _newWorkHistory["new_casestatusreason"] = new EntityReference(CaseStatus.LogicalName, CaseStatus.Id);
            }//
            _newWorkHistory["statuscode"] = new OptionSetValue(1);//1 - active
          // _newWorkHistory["new_timespendbeforerouting"] = _preWorkHistory["new_timespendbeforerouting"];
            _newWorkHistory["new_case"] = Case.ToEntityReference();
            _newWorkHistory["new_name"] = Case["title"];
            if(QueueItem != null)
            {
                _newWorkHistory["new_queue"] = new EntityReference(QueueItem.GetAttributeValue<EntityReference>("queueid").LogicalName, QueueItem.GetAttributeValue<EntityReference>("queueid").Id);
                _newWorkHistory["new_queueitem"] = QueueItem.ToEntityReference();
                _newWorkHistory["new_enteredqueue"] = QueueItem.GetAttributeValue<DateTime>("enteredon");
                Owner = FetchOwner(QueueItem);
            }
            else if(_preWorkHistory.Attributes.Contains("new_queue"))
            {
                _newWorkHistory["new_queue"] = _preWorkHistory["new_queue"];
                _newWorkHistory["new_queueitem"] = _preWorkHistory["new_queueitem"];
                _newWorkHistory["new_enteredqueue"] = _preWorkHistory["new_enteredqueue"];
            }
           
            if (Owner != null)
            {
                if (Owner.LogicalName.ToLower().Equals("systemuser"))
                {
                    _newWorkHistory["new_workedbyuser"] = new EntityReference(Owner.LogicalName, Owner.Id);
                }
                else
                {
                    _newWorkHistory["new_workedbyteam"] = new EntityReference(Owner.LogicalName, Owner.Id);
                }
            }

            Guid _newID = service.Create(_newWorkHistory);
            return _newID;
        }
        public bool UpdateWorkHistory(Entity _preWorkHistory, IOrganizationService service, Entity Case, ITracingService tracing)
        {
            tracing.Trace("inside UpdateWorkHistory method");
            bool isUpdated = false;
            //inactivate the work history record
            _preWorkHistory["new_timespendbycurrentuser"] = _preWorkHistory.Attributes.Contains("new_calctimespendbycurruser") ? _preWorkHistory.GetAttributeValue<decimal>("new_calctimespendbycurruser") : new decimal(0.00);
            _preWorkHistory["new_timespendbycaseinqueue"] = _preWorkHistory.Attributes.Contains("new_calctimespendqueue") ? _preWorkHistory.GetAttributeValue<decimal>("new_calctimespendqueue") : new decimal(0.00);
            _preWorkHistory["new_totaltimespendoncase"] = _preWorkHistory.Attributes.Contains("new_calctotaltimespendoncase") ? _preWorkHistory.GetAttributeValue<decimal>("new_calctotaltimespendoncase") : new decimal(0.00);
            if (!_preWorkHistory.Attributes.Contains("new_workedbyuser") && !_preWorkHistory.Attributes.Contains("new_workedbyteam"))
            {
                DateTime _caseModifiedon = Case.GetAttributeValue<DateTime>("modifiedon");
                DateTime _workHistoryCreatedon = _preWorkHistory.GetAttributeValue<DateTime>("createdon");
                TimeSpan _timebeforeRouting = _caseModifiedon.Subtract(_workHistoryCreatedon);
                    
                _preWorkHistory["new_timespendbeforerouting"] = Convert.ToDecimal(_timebeforeRouting.TotalMinutes);
            }
            
            EntityReference CaseStatus = FetchCaseStatus(Case);
            if (CaseStatus != null)
            {
                _preWorkHistory["new_casestatusreason"] = new EntityReference(CaseStatus.LogicalName, CaseStatus.Id);
            }
            _preWorkHistory["statuscode"] = new OptionSetValue(100000000);//closed-100000000
            service.Update(_preWorkHistory);
            isUpdated = true;
            tracing.Trace("Updated existing  Work history:" + _preWorkHistory.Id.ToString());
            return isUpdated;
        }
        public Entity FetchPreviousWorkHistory(IOrganizationService service, Guid CaseId, ITracingService tracing)
        {
            string FetchWorkHistory = $@"<fetch >
                                          <entity name='new_queueworkhistory' >
                                            <attribute name='new_workedbyteam' />
                                            <attribute name='new_case' />
                                            <attribute name='new_name' />
                                            <attribute name='new_calculatedtimespendbycaseinqueue' />
                                            <attribute name='new_calctotaltimespendoncase' />
                                            <attribute name='new_calctimespendbycurruser' />
                                            <attribute name='new_casestatusreason' />
                                            <attribute name='createdon' />
                                            <attribute name='new_workedbyuser' />                                            
                                            <attribute name='new_timespendbeforerouting' />
                                            <attribute name='new_queue' />
                                            <attribute name='new_queueitem' />
                                            <attribute name='new_enteredqueue' />
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

        public Entity RecordFetch(IOrganizationService service, string EntityLogicalName, Guid ID, string[] Columnset)
        {
            Entity RecordTobeFetched = service.Retrieve(EntityLogicalName, ID, new ColumnSet(Columnset));
            return RecordTobeFetched;
        }

        public EntityReference FetchCaseStatus(Entity Target)
        {
            EntityReference CaseStatus = null;

            if (Target.Attributes.Contains("cr32a_casestatusreason"))
            {
                CaseStatus = new EntityReference(Target.GetAttributeValue<EntityReference>("cr32a_casestatusreason").LogicalName, Target.GetAttributeValue<EntityReference>("cr32a_casestatusreason").Id);
            }

            return CaseStatus;

        }

        public EntityReference FetchOwner(Entity Target)
        {
            EntityReference Owner = null;
            if (Target.LogicalName.ToUpper().Equals(Modal.CaseLogicalName.ToUpper()))
            {
                if (Target.Attributes.Contains("ownerid"))
                {
                    Owner = new EntityReference(Target.GetAttributeValue<EntityReference>("ownerid").LogicalName, Target.GetAttributeValue<EntityReference>("ownerid").Id);
                }
            }
            else if(Target.LogicalName.ToUpper().Equals(Modal.QueueITemLogicalName.ToUpper()))
            {
                if (Target.Attributes.Contains("workerid"))
                {
                    Owner = new EntityReference(Target.GetAttributeValue<EntityReference>("workerid").LogicalName, Target.GetAttributeValue<EntityReference>("workerid").Id);
                }
            }

            return Owner;

        }
    }
}
