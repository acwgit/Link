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
    public class DealGetOwnerBuName : CodeActivity
    {


        [RequiredArgument]
        [ReferenceTarget("clf_deal")]
        [Input("Deal")]
        public InArgument<EntityReference> _dealRef { get; set; }


        [Output("Bu Name")]
        public OutArgument<string> _output { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.InitiatingUserId);
            ITracingService tracer = executionContext.GetExtension<ITracingService>();


            var dealRef = this._dealRef.Get(executionContext);

            try
            {
                Entity dealEn = service.Retrieve(dealRef.LogicalName, dealRef.Id, new ColumnSet("ownerid"));
                EntityReference ownerRef = dealEn.GetAttributeValue<EntityReference>("ownerid");

                tracer.Trace("Deal ID: {0}", dealEn.Id.ToString());

                if (ownerRef == null) 
                {
                    tracer.Trace("Owner Not Found");
                    return;
                }

                Entity ownerEn = service.Retrieve(ownerRef.LogicalName, ownerRef.Id, new ColumnSet("businessunitid"));
                EntityReference buRef = ownerEn.GetAttributeValue<EntityReference>("businessunitid");

                if (buRef == null)
                {
                    tracer.Trace("Bu Not Found");
                    return;
                }

                Entity buEn = service.Retrieve(buRef.LogicalName, buRef.Id, new ColumnSet("name"));
                string buName = buEn.GetAttributeValue<string>("name");

                tracer.Trace("Deal Owner Bu Name: {0}", buName);

                this._output.Set(executionContext, buName);
                
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

