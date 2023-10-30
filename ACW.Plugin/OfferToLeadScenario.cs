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
    public class OfferToLeadScenario : CodeActivity
    {


        [RequiredArgument]
        [ReferenceTarget("clf_leadscenario")]
        [Input("Lead Scenario")]
        public InArgument<EntityReference> _leadScenarioRef { get; set; }

        [RequiredArgument]        
        [Input("Target Field Logical Name")]
        public InArgument<string> _fieldName { get; set; }        

        [RequiredArgument]
        [Input("Field Type (string/money)")]
        public InArgument<string> _fieldType { get; set; }

        
        [Output("String Output")]
        public OutArgument<string> _stringOutput { get; set; }

        [Output("Money Output")]
        public OutArgument<Money> _moneyOutput { get; set; }     



        protected override void Execute(CodeActivityContext executionContext)
        {
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.InitiatingUserId);
            ITracingService tracer = executionContext.GetExtension<ITracingService>();

            


            var leadScenarioRef = this._leadScenarioRef.Get(executionContext);
            var fieldName = this._fieldName.Get(executionContext);
            var fieldType = this._fieldType.Get(executionContext);

            try
            {
                if (leadScenarioRef == null) 
                {
                    tracer.Trace("No Lead Scenario");
                    return;
                }

                tracer.Trace("Lead Scenario ID: {0}", leadScenarioRef.Id.ToString());
                tracer.Trace("Target Field Logical Name: {0}", fieldName);
                tracer.Trace("Field Type: {0}", fieldType);

                QueryExpression offerQe = new QueryExpression("clf_offer");
                offerQe.Criteria.AddCondition("lms_leadscenario", ConditionOperator.Equal, leadScenarioRef.Id);
                offerQe.Criteria.AddCondition("clf_offertype", ConditionOperator.Equal, 100000000); // Asking Rent

                offerQe.ColumnSet.AddColumns(fieldName);

                EntityCollection offerEc = service.RetrieveMultiple(offerQe);

                Entity offerEn = offerEc.Entities.FirstOrDefault();

                if (offerEn == null) 
                {
                    tracer.Trace("No Offer Found");
                }

                switch(fieldType) 
                {
                    case "string":
                        {
                            var stringOutput_ = offerEn.GetAttributeValue<string>(fieldName);
                            this._stringOutput.Set(executionContext, stringOutput_);
                        }
                        break;
                    case "money":
                        {
                            var moneyOutput_ = offerEn.GetAttributeValue<Money>(fieldName);
                            this._moneyOutput.Set(executionContext, moneyOutput_);
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
    }
}

