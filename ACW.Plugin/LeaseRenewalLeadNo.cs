using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace ACW.Plugin
{
    public class LeaseRenewalLeadNo : CodeActivity
    {
        [RequiredArgument]
        [Input("Lease")]
        [ReferenceTarget("lms_lease")]
        public InArgument<EntityReference> _lease { get; set; }

        [Output("Renewal Lead Number")]
        public OutArgument<string> _output { get; set; }

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
                Entity lease = service.Retrieve("lms_lease", this._lease.Get(executionContext).Id, new ColumnSet("lms_leasenumber", "lms_revisionid", "lms_leadgenerationseqno"));
                string leaseNumber = lease.GetAttributeValue<string>("lms_leasenumber");
                int revisionId = lease.GetAttributeValue<int>("lms_revisionid");
                int seq = lease.GetAttributeValue<int>("lms_leadgenerationseqno");

                string newNumber = "L" + leaseNumber.Replace(leaseNumber[0].ToString(), "") + "/" + string.Format("{0:0000}", revisionId) + "/" + seq.ToString();
                this._output.Set(executionContext, newNumber);

            }
            catch (Exception ex)
            {
                string msg = $"{this.GetType().Name} of Error Message: \r\n \t {ex.Message}";
                tracer.Trace(msg);
                throw new InvalidPluginExecutionException(msg);
            }
        }
    }
}
