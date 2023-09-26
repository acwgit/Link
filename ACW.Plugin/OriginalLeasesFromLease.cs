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
    public class OriginalLeasesFromLease : CodeActivity
    {

        [RequiredArgument]
        [Input("Lease ID")]        
        public InArgument<string> _leaseId { get; set; }

        [RequiredArgument]
        [Input("Lead")]
        [ReferenceTarget("clf_lead")]
        public InArgument<EntityReference> _leadRef { get; set; }
        protected override void Execute(CodeActivityContext executionContext)
        {
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.InitiatingUserId);
            ITracingService tracer = executionContext.GetExtension<ITracingService>();


            try
            {
                string leaseId = this._leaseId.Get(executionContext);
                EntityReference leadRef = this._leadRef.Get(executionContext);
                
                if (leadRef == null)
                {
                    tracer.Trace("No Lead Found");
                    return;
                }

                QueryExpression leaseQe = new QueryExpression("lms_lease");
                leaseQe.Criteria.AddCondition("lms_leasenumber", ConditionOperator.Equal, leaseId);
                leaseQe.Criteria.AddCondition("lms_activatedlease", ConditionOperator.Equal, true);
                leaseQe.ColumnSet.AddColumns("lms_lead");

                EntityCollection leaseEc = service.RetrieveMultiple(leaseQe);

                foreach (Entity lease in leaseEc.Entities)
                {
                    Entity new_originalLeaseEn = new Entity("lms_leadlease");
                    new_originalLeaseEn["lms_lead"] = leadRef;
                    new_originalLeaseEn["lms_lease"] = new EntityReference(lease.LogicalName,lease.Id);                    

                    var new_guid = service.Create(new_originalLeaseEn);

                    tracer.Trace("Original Lease Created With Id: {0}", new_guid.ToString());
                }

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

