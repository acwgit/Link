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
    public class LeadUnitFromLease : CodeActivity
    {
       
        [RequiredArgument]
        [Input("Lease")]
        [ReferenceTarget("lms_lease")]
        public InArgument<EntityReference> _leaseRef { get; set; }

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
                EntityReference leaseRef = this._leaseRef.Get(executionContext);
                EntityReference leadRef = this._leadRef.Get(executionContext);
                if (leaseRef == null) 
                {
                    tracer.Trace("No Lease Found");
                    return;
                }

                if (leadRef == null)
                {
                    tracer.Trace("No Lead Found");
                    return;
                }

                Entity leaseEn = service.Retrieve(leaseRef.LogicalName, leaseRef.Id, new ColumnSet("lms_property"));
                EntityReference propertyRef = leaseEn.GetAttributeValue<EntityReference>("lms_property");

                QueryExpression leaseUnitQe = new QueryExpression("lms_leaseunit");
                leaseUnitQe.Criteria.AddCondition("lms_lease", ConditionOperator.Equal, leaseRef.Id);
                leaseUnitQe.ColumnSet.AddColumns("lms_unit", "lms_lease");

                EntityCollection leaseUnitEc = service.RetrieveMultiple(leaseUnitQe);

                foreach(Entity leaseUnit in leaseUnitEc.Entities) 
                {
                    Entity new_leadUnitEn = new Entity("lms_leadunit");
                    new_leadUnitEn["lms_lead"] = leadRef;
                    new_leadUnitEn["lms_property"] = propertyRef;
                    new_leadUnitEn["lms_unit"] = leaseUnit.GetAttributeValue<EntityReference>("lms_unit");

                    var new_guid = service.Create(new_leadUnitEn);

                    tracer.Trace("Lead Unit Created With Id: {0}", new_guid.ToString());
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

