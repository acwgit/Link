using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;

namespace ACW.Plugin
{
    public class CreateBatchRecordAndMapToTask : CodeActivity
    {


        [RequiredArgument]
        [ReferenceTarget("task")]
        [Input("Task")]
        public InArgument<EntityReference> _taskRef { get; set; }       

        protected override void Execute(CodeActivityContext executionContext)
        {
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.InitiatingUserId);
            ITracingService tracer = executionContext.GetExtension<ITracingService>();


            var taskRef = this._taskRef.Get(executionContext);

            try
            {
                Entity taskEn = service.Retrieve(taskRef.LogicalName, taskRef.Id, new ColumnSet("lms_taskcategory", "lms_batchkey"));
                int taskCategory = taskEn.Contains("lms_taskcategory") ? taskEn.GetAttributeValue<OptionSetValue>("lms_taskcategory").Value : 0;
                string batchKey = taskEn.GetAttributeValue<string>("lms_batchkey");

                tracer.Trace("Task ID: {0}", taskEn.Id.ToString());
                tracer.Trace("taskCategory : {0}", taskCategory.ToString());
                tracer.Trace("batchKey: {0}", batchKey);

                switch (taskCategory) 
                {
                    case 176660000: // Legal
                        {
                            MainProcess(0, taskEn, "lms_legalbatch", "lms_batchkey", batchKey, service, tracer);
                        }
                        break;

                    case 176660001: // Finance
                        {
                            MainProcess(0, taskEn, "lms_financebatch", "lms_batchkey", batchKey, service, tracer);   
                        }
                        break;

                    case 176660003: // Tenancy
                        {
                            MainProcess(0, taskEn, "lms_tenancyadminbatch", "lms_batchkey", batchKey, service, tracer);
                        }
                        break;
                }




            }
            catch (Exception ex)
            {
                string msg = $"{this.GetType().Name} of Error Message: \r\n \t {ex.Message}";
                tracer.Trace(msg);
                throw new InvalidPluginExecutionException(msg);
            }
        }

        private Entity FindBatchRecord(string entityLogicalName, string targetFieldLogicalName, string batchKey, IOrganizationService service) 
        {
            QueryExpression batchRecordQe = new QueryExpression(entityLogicalName);
            batchRecordQe.Criteria.AddCondition(targetFieldLogicalName, ConditionOperator.Equal, batchKey);

            EntityCollection batchRecordEc = service.RetrieveMultiple(batchRecordQe);

            return batchRecordEc.Entities.FirstOrDefault();
        }

        private EntityReference CreateBatchRecord(string entityLogicalName, string batchKeyFieldLogicalName, string batchKey, IOrganizationService service) 
        {
            Entity new_BatchEn = new Entity(entityLogicalName);
            new_BatchEn[batchKeyFieldLogicalName] = batchKey;

            var new_Guid = service.Create(new_BatchEn);
            
            return new EntityReference(entityLogicalName, new_Guid);
        }

        private void MainProcess(int retryCount,Entity taskEn, string entityLogicalName, string targetFieldLogicalName, string batchKey, IOrganizationService service, ITracingService tracer) 
        {
            if (retryCount == 5) 
            {
                return;
            }
            try
            {
                Entity legalBatchRecord = FindBatchRecord(entityLogicalName, targetFieldLogicalName, batchKey, service);
                EntityReference new_BatchRecordRef = null;
                if (legalBatchRecord == null)
                {
                    Entity new_BatchEn = new Entity(entityLogicalName);
                    new_BatchEn[targetFieldLogicalName] = batchKey;

                    var new_Guid = service.Create(new_BatchEn);

                    new_BatchRecordRef = new EntityReference(entityLogicalName, new_Guid);
                }
                else
                {
                    new_BatchRecordRef = legalBatchRecord.ToEntityReference();
                }

                Entity batchRecord = service.Retrieve(new_BatchRecordRef.LogicalName, new_BatchRecordRef.Id, new ColumnSet("lms_name"));
                string batchNo = batchRecord.GetAttributeValue<string>("lms_name");

                Entity update_BatchRecordEn = new Entity(taskEn.LogicalName, taskEn.Id);
                update_BatchRecordEn[entityLogicalName] = new_BatchRecordRef;
                update_BatchRecordEn["lms_batchnolegalstamping"] = batchNo;

                service.Update(update_BatchRecordEn);
                tracer.Trace("Updated Batch Record ID: {0}", new_BatchRecordRef.Id.ToString());
            }
            catch 
            {
                tracer.Trace("Retrying count {0}", (retryCount + 1).ToString());
                MainProcess(retryCount + 1, taskEn, entityLogicalName, targetFieldLogicalName, batchKey, service, tracer);
            }
        }
    }
}

