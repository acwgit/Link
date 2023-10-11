using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace ACW.Plugin
{
    public class AccountContactPvBillingSyncUpdate : CodeActivity
    {
        [RequiredArgument]
        [Input("Account")]
        [ReferenceTarget("account")]
        public InArgument<EntityReference> _account { get; set; }

        [RequiredArgument]
        [Input("Role Description")]
        public InArgument<string> _role { get; set; }

        IOrganizationService service;
        ITracingService tracer;

        protected override void Execute(CodeActivityContext executionContext)
        {
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            this.service = serviceFactory.CreateOrganizationService(context.InitiatingUserId);
            this.tracer = executionContext.GetExtension<ITracingService>();

            try
            {
                EntityReference account = this._account.Get(executionContext);
                Entity role = GetRole(this._role.Get(executionContext));
                tracer.Trace("Account: {0}", account.ToString());
                tracer.Trace("Role: {0}", role.ToString());

                Entity contact = GetContact(account, role);
                if (contact != null)
                {
                    Entity updateEn = new Entity(contact.LogicalName, contact.Id);
                    updateEn["lms_allowtosynctoyardi"] = true;
                    service.Update(updateEn);
                }
            }
            catch (Exception ex)
            {
                string msg = $"{this.GetType().Name} of Error Message: \r\n \t {ex.Message}";
                tracer.Trace(msg);
                throw new InvalidPluginExecutionException(msg);
            }
        }

        private Entity GetContact(EntityReference account, Entity role)
        { 
            QueryExpression qe = new QueryExpression("contact");
            qe.Criteria.AddCondition("lms_account", ConditionOperator.Equal, account.Id);
            qe.Criteria.AddCondition("lms_roledescription", ConditionOperator.Equal, role.Id);
            qe.Criteria.AddCondition("lms_allowtosynctoyardi", ConditionOperator.Equal, false);
            qe.ColumnSet.AddColumn("lms_allowtosynctoyardi");

            EntityCollection ec = service.RetrieveMultiple(qe);
            int count = ec.Entities.Count;

            if (count > 0)
            {
                if (count == 1)
                    return ec.Entities[0];
                else
                    throw new InvalidPluginExecutionException("More than one Contact has been found");
            }

            else
                return null;         
        }

        private Entity GetRole(string role)
        {
            QueryExpression qe = new QueryExpression("lms_roledescription");
            qe.Criteria.AddCondition("lms_name", ConditionOperator.Equal, role);

            EntityCollection ec = service.RetrieveMultiple(qe);

            if (ec.Entities.Count == 0)
                throw new InvalidPluginExecutionException("Unable to find the Role: " + role);
            else
                return ec.Entities[0];
        }
    }
}
