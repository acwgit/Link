﻿using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace ACW.Plugin
{
    public class AccountContactPvBillingSyncOutput : CodeActivity
    {
        [RequiredArgument]
        [Input("Account")]
        [ReferenceTarget("account")]
        public InArgument<EntityReference> _account { get; set; }

        [RequiredArgument]
        [Input("Role Description")]
        public InArgument<string> _role { get; set; }

        [RequiredArgument]
        [Input("Field Name")]
        public InArgument<string> _fieldName { get; set; }

        [Output("Done or not")]
        public OutArgument<bool> _done { get; set; }

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
                string field = this._fieldName.Get(executionContext);
                tracer.Trace("Account: {0}", account.Id.ToString());
                tracer.Trace("Role: {0}", role.Id.ToString());
                tracer.Trace("Field: {0}", field);

                Entity contact = GetContact(account, role, field);
                if (contact != null)
                {
                    var done = contact.GetAttributeValue<bool>(field);
                    tracer.Trace("Done: {0}", done.ToString());
                    this._done.Set(executionContext, done);
                }
            }
            catch (Exception ex)
            {
                string msg = $"{this.GetType().Name} of Error Message: \r\n \t {ex.Message}";
                tracer.Trace(msg);
                throw new InvalidPluginExecutionException(msg);
            }
        }

        private Entity GetContact(EntityReference account, Entity role, string field)
        {
            QueryExpression qe = new QueryExpression("contact");
            qe.Criteria.AddCondition("lms_account", ConditionOperator.Equal, account.Id);
            qe.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            qe.Criteria.AddCondition("lms_roledescription", ConditionOperator.Equal, role.Id);
            qe.ColumnSet.AddColumn(field);

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
