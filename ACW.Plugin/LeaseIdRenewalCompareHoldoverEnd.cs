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
    public class LeaseIdRenewalCompareHoldoverEnd : CodeActivity
    {
        [RequiredArgument]
        [Input("Lease ID")]
        public InArgument<string> _leaseId { get; set; }

        [RequiredArgument]
        [Input("Holdover End")]
        public InArgument<DateTime> _holdoverEnd { get; set; }

        [Output("Output")]
        public OutArgument<bool> _output { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.InitiatingUserId);
            ITracingService tracer = executionContext.GetExtension<ITracingService>();

            try
            {
                string leaseId = this._leaseId.Get(executionContext);
                DateTime holdoverEnd = this._holdoverEnd.Get(executionContext);
                

                if (leaseId == "")
                    throw new InvalidPluginExecutionException("Lease ID is empty.");

                if (holdoverEnd.Year == 1)
                    throw new InvalidPluginExecutionException("Holdover End Datetime is empty");

                tracer.Trace("Leasd Id: {0}", leaseId);
                tracer.Trace("HoldOver End: {0}", holdoverEnd.ToString());

                EntityCollection ec = GetLease(service, leaseId, holdoverEnd);
                tracer.Trace("Count: {0}", ec.Entities.Count.ToString());

                if (ec.Entities.Count >= 1)
                    this._output.Set(executionContext, true);
                else
                    this._output.Set(executionContext, false);
            }
            catch (Exception ex)
            {
                string msg = $"{this.GetType().Name} of Error Message: \r\n \t {ex.Message}";
                tracer.Trace(msg);
                throw new InvalidPluginExecutionException(msg);
            }
        }

        private EntityCollection GetLease(IOrganizationService service, string leaseId, DateTime holdoverEnd)
        {
            QueryExpression qe = new QueryExpression("lms_lease");
            qe.Criteria.AddCondition("lms_leasenumber", ConditionOperator.Equal, leaseId);
            qe.Criteria.AddCondition("lms_leasetype", ConditionOperator.Equal, 1);
            qe.Criteria.AddCondition("lms_leasestart", ConditionOperator.OnOrBefore, holdoverEnd);
            qe.Criteria.AddCondition("lms_leaseend", ConditionOperator.OnOrAfter, holdoverEnd);

            EntityCollection ec = service.RetrieveMultiple(qe);

            return ec;
        }
    }
}

