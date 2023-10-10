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
    public class AccountCheckContactRoleDescription : CodeActivity
    {


        [RequiredArgument]
        [ReferenceTarget("account")]
        [Input("Account")]
        public InArgument<EntityReference> _accountRef { get; set; }

        [RequiredArgument]        
        [Input("Role Description Text")]
        public InArgument<string> _roleDesText { get; set; }


        [Output("Exist")]
        public OutArgument<bool> _output { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.InitiatingUserId);
            ITracingService tracer = executionContext.GetExtension<ITracingService>();


            var accountRef = this._accountRef.Get(executionContext);
            var roleDesText = this._roleDesText.Get(executionContext);
            var result = false;

            try
            {
                tracer.Trace("Account ID: {0}", accountRef.Id.ToString());
                tracer.Trace("roleDesText: {0}", roleDesText);

                QueryExpression contactQe = new QueryExpression("contact");
                contactQe.Criteria.AddCondition("lms_account", ConditionOperator.Equal, accountRef.Id);                
                contactQe.LinkEntities.Add(new LinkEntity("contact", "lms_roledescription", "lms_roledescription", "lms_roledescriptionid", JoinOperator.Inner));
                contactQe.LinkEntities[0].EntityAlias = "roledescription";
                contactQe.LinkEntities[0].Columns.AddColumns("lms_name");

                EntityCollection contactEc = service.RetrieveMultiple(contactQe);

                foreach (Entity contactEn in contactEc.Entities) 
                {
                    string roleDescription = (string)contactEn.GetAttributeValue<AliasedValue>("roledescription.lms_name").Value;
                    tracer.Trace("roleDescription: {0}", roleDescription);

                    if (roleDescription == roleDesText) 
                    {
                        result = true;
                    }
                }

                tracer.Trace("result: {0}", result.ToString());

                this._output.Set(executionContext, result);

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

