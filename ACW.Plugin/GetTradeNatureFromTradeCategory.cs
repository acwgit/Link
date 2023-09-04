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
    public class GetTradeNatureFromTradeCategory : CodeActivity
    {
        
        [RequiredArgument]
        [ReferenceTarget("lms_tradecategory")]
        [Input("Trade Category")]
        public InArgument<EntityReference> _tradeCategoryRef { get; set; }
        [RequiredArgument]
        [Input("LookUp Logical Name")]
        public InArgument<string> _lookupLogicalName { get; set; }

        [ReferenceTarget("lms_tradenature")]
        [Output("Trade Nature")]
        public OutArgument<EntityReference> _output { get; set; }
        //string err = string.Empty;
        protected override void Execute(CodeActivityContext executionContext)
        {
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.InitiatingUserId);
            ITracingService tracer = executionContext.GetExtension<ITracingService>();


            var lookupLogicalName = this._lookupLogicalName.Get(executionContext);


            try
            {
                EntityReference tradeCategoryRef = this._tradeCategoryRef.Get(executionContext);
                if (tradeCategoryRef == null) 
                {
                    return;
                }

                Entity tradeCategoryEn = service.Retrieve(tradeCategoryRef.LogicalName, tradeCategoryRef.Id, new ColumnSet(lookupLogicalName));

                EntityReference tradeNatureRef = tradeCategoryEn.GetAttributeValue<EntityReference>(lookupLogicalName);

                if (tradeNatureRef != null)
                {
                    this._output.Set(executionContext, new EntityReference(tradeNatureRef.LogicalName, tradeNatureRef.Id));
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

