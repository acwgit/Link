using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace ACW.Plugin
{
    public class LeadScenarioUpdateAskingRentOffer : CodeActivity
    {
        [RequiredArgument]
        [Input("Lead Scenario")]
        [ReferenceTarget("clf_leadscenario")]
        public InArgument<EntityReference> _leadSc { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.InitiatingUserId);
            ITracingService tracer = executionContext.GetExtension<ITracingService>();

            try
            {
                EntityReference leadSc = this._leadSc.Get(executionContext);
                tracer.Trace("Lead Scenario: {0}", leadSc.Id.ToString());

                QueryExpression qe = new QueryExpression("clf_offer");
                qe.Criteria.AddCondition("lms_leadscenario", ConditionOperator.Equal, leadSc.Id);
                qe.Criteria.AddCondition("clf_offertype", ConditionOperator.Equal, 100000000);
                qe.ColumnSet.AddColumn("statuscode");

                EntityCollection ec = service.RetrieveMultiple(qe);

                if (ec.Entities.Count > 0)
                {
                    foreach (var offer in ec.Entities)
                    { 
                        Entity updateEn = new Entity(offer.LogicalName, offer.Id);
                        updateEn["statuscode"] = new OptionSetValue(176660005);
                        service.Update(updateEn);

                        tracer.Trace("Updated offer: {0}", offer.Id.ToString());
                    }
                }

                else
                    tracer.Trace("No offer");
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

