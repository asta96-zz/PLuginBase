using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;

namespace Retro.Plugins.Common
{
    public class BusinessLogic
    {
        private IOrganizationService AdminService { get; set; }
        private ITracingService TracingService { get; set; }

        public BusinessLogic() { }
        public BusinessLogic(IOrganizationServiceFactory serviceFactory, ITracingService tracing)
        {
            AdminService = serviceFactory.CreateOrganizationService(null);
            TracingService = tracing;
        }
        public Guid CreateWorkHistory(Entity _preWorkHistory, Entity Case, IOrganizationService service, Entity QueueItem = null)
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
            if (QueueItem != null)
            {
                _newWorkHistory["new_queue"] = new EntityReference(QueueItem.GetAttributeValue<EntityReference>("queueid").LogicalName, QueueItem.GetAttributeValue<EntityReference>("queueid").Id);
                _newWorkHistory["new_queueitem"] = QueueItem.ToEntityReference();
                _newWorkHistory["new_enteredqueue"] = QueueItem.GetAttributeValue<DateTime>("enteredon");
                Owner = FetchOwner(QueueItem);
            }
            else if ((_preWorkHistory != null) && _preWorkHistory.Attributes.Contains("new_queue"))
            {
                _newWorkHistory["new_queue"] = _preWorkHistory["new_queue"];
                _newWorkHistory["new_queueitem"] = _preWorkHistory["new_queueitem"];
                _newWorkHistory["new_enteredqueue"] = _preWorkHistory["new_enteredqueue"];
            }

            if (Owner != null)
            {
                if (string.Equals(Owner.LogicalName, "systemuser", StringComparison.OrdinalIgnoreCase))
                {
                    _newWorkHistory["new_workedbyuser"] = new EntityReference(Owner.LogicalName, Owner.Id);
                }
                else
                {
                    TracingService.Trace("before fetching queueItem ");
                    Entity queueItem = FetchQueueItem(Case.Id);
                    TracingService.Trace("after fetching queueItem {0}:queueid:{1} enteredQueue:{2}",queueItem.Id,queueItem.GetAttributeValue<EntityReference>("queueid").Id,queueItem.GetAttributeValue<DateTime>("enteredon"));
                    _newWorkHistory["new_queue"] = queueItem["queueid"];
                    _newWorkHistory["new_queueitem"] = queueItem.ToEntityReference();
                    _newWorkHistory["new_enteredqueue"] = queueItem["enteredon"];
                }
            }
            Guid _newID = AdminService.Create(_newWorkHistory);
            return _newID;
        }

        private Entity FetchQueueItem(Guid caseId)
        {
            TracingService.Trace("entering FetchQueueItem ");
            Entity queueItem = null;
            QueryExpression QEqueueItem = new QueryExpression("queueitem");
            QEqueueItem.Criteria.AddCondition(new ConditionExpression("objectid", ConditionOperator.Equal, caseId));
            QEqueueItem.ColumnSet.AddColumns("queueid", "enteredon");
            EntityCollection queueItems = AdminService.RetrieveMultiple(QEqueueItem);
            if(queueItems!=null)
            {
                queueItem = queueItems.Entities.FirstOrDefault();
            }
            TracingService.Trace("exiting FetchQueueItem ");
            return queueItem;
        }

        internal string FetchQueueName(Guid queueId)
        {
            var queue = AdminService.Retrieve("queue", queueId, new ColumnSet("name"));
            return (queue != null ? queue["name"].ToString() : "");
        }

        internal void PreCreateWH(Entity caseContext)
        {
            Entity workHistory = new Entity("new_queueworkhistory");
            EntityReference CaseStatus = FetchCaseStatus(caseContext);
            EntityReference Owner = FetchOwner(caseContext);
            workHistory.Attributes.Add("new_name", caseContext["title"]);
            workHistory.Attributes.Add("new_case", caseContext.ToEntityReference());
            workHistory.Attributes.Add("new_casestatusreason", CaseStatus);
            workHistory.Attributes.Add("statuscode", new OptionSetValue(1));
            if (string.Equals(Owner.LogicalName, "systemuser", StringComparison.OrdinalIgnoreCase))
            {
                workHistory.Attributes.Add("new_workedbyuser", Owner);
            }
            Guid workHistoryId = AdminService.Create(workHistory);
            TracingService.Trace("New Work history created with ID {0} at UTC: {1}", workHistoryId, DateTime.UtcNow.ToString(@"MM\/dd\/yyyy HH:mm"));
        }

        internal void updateCaseOwner(Entity caseRecord, string queueName)
        {
            EntityReference owner = FetchTeamFromQueue(queueName);
            if (owner != null)
            {
                caseRecord["ownerid"]= owner;
                TracingService.Trace("before updating Case with new owner {0}", owner.Id);
                AdminService.Update(caseRecord);
            }
        }

        private EntityReference FetchTeamFromQueue(string queueName)
        {
            EntityReference owner = null;
            QueryExpression QEteam = new QueryExpression("team");
            QEteam.Criteria.AddCondition(new ConditionExpression("name", ConditionOperator.Equal, queueName));
            QEteam.ColumnSet.AddColumn("teamid");
            var teams = AdminService.RetrieveMultiple(QEteam);
            if (teams != null && teams.Entities.Count > 0)
            {
                var team = teams.Entities.FirstOrDefault();
                owner = team.ToEntityReference();
            }
            return owner;
        }

        public bool UpdateWorkHistory(Entity _preWorkHistory, Entity Case, ITracingService tracing)
        {
            tracing.Trace("inside UpdateWorkHistory method");
            bool isUpdated = false;
            //inactivate the work history record
            _preWorkHistory["new_timespendbycurrentuser"] = _preWorkHistory.Attributes.Contains("new_calctimespendbycurruser") ? _preWorkHistory.GetAttributeValue<decimal>("new_calctimespendbycurruser") : new decimal(0.00);
            _preWorkHistory["new_timespendbycaseinqueue"] = _preWorkHistory.Attributes.Contains("new_calctimespendqueue") ? _preWorkHistory.GetAttributeValue<decimal>("new_calctimespendqueue") : new decimal(0.00);
            _preWorkHistory["new_totaltimespendoncase"] = _preWorkHistory.Attributes.Contains("new_calctotaltimespendoncase") ? _preWorkHistory.GetAttributeValue<decimal>("new_calctotaltimespendoncase") : new decimal(0.00);
            if (!_preWorkHistory.Attributes.Contains("new_workedbyuser"))
            {
                DateTime _caseModifiedon = Case.GetAttributeValue<DateTime>("modifiedon");
                DateTime _workHistoryCreatedon = _preWorkHistory.GetAttributeValue<DateTime>("createdon");
                TimeSpan _timebeforeRouting = _caseModifiedon.Subtract(_workHistoryCreatedon);
                _preWorkHistory["new_timespendbeforerouting"] = Convert.ToDecimal(_timebeforeRouting.TotalMinutes);
            }
            EntityReference CaseStatus = FetchCaseStatus(Case);
            if (CaseStatus != null)
            {
                _preWorkHistory["new_casestatusreason"] = CaseStatus;
            }
            _preWorkHistory["statuscode"] = new OptionSetValue(100000000);//closed-100000000
            AdminService.Update(_preWorkHistory);
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
            EntityCollection coll = AdminService.RetrieveMultiple(new FetchExpression(FetchWorkHistory));

            if (coll.Entities.Count > 0)
            {
                tracing.Trace("Fetched WorkHistory Count:" + coll.Entities.Count);
                History = coll.Entities.FirstOrDefault();
            }
            else
            {
                FetchWorkHistory = $@"<fetch >
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
                                            <filter type='and' >
                                              <condition attribute='new_case' operator='eq' value='{CaseId}' />
                                              <condition attribute='new_timespendbeforerouting' operator='not-null' />
                                              </filter>
                                          </entity>
                                        </fetch>";
                coll = AdminService.RetrieveMultiple(new FetchExpression(FetchWorkHistory));

                if (coll.Entities.Count > 0)
                {
                    tracing.Trace("Fetched WorkHistory Count:" + coll.Entities.Count);
                    History = coll.Entities.FirstOrDefault();
                }
            }

            return History;
        }

        public Entity RecordFetch(IOrganizationService service, string EntityLogicalName, Guid ID, string[] Columnset)
        {
            Entity RecordTobeFetched = AdminService.Retrieve(EntityLogicalName, ID, new ColumnSet(Columnset));
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
            else if (Target.LogicalName.ToUpper().Equals(Modal.QueueITemLogicalName.ToUpper()))
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