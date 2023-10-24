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
    public class PropertyGetUser : CodeActivity
    {


        [RequiredArgument]
        [ReferenceTarget("clf_property")]
        [Input("Property")]
        public InArgument<EntityReference> _propertyRef { get; set; }

        [RequiredArgument]
        [ReferenceTarget("businessunit")]
        [Input("BU")]
        public InArgument<EntityReference> _BURef { get; set; }

        [RequiredArgument]        
        [Input("Role (Approver/CD Approver/GM/AGM)")]
        public InArgument<string> _roleText { get; set; }


        [ReferenceTarget("systemuser")]
        [Output("User")]
        public OutArgument<EntityReference> _output { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.InitiatingUserId);
            ITracingService tracer = executionContext.GetExtension<ITracingService>();


            var propertyRef = this._propertyRef.Get(executionContext);
            var BURef = this._BURef.Get(executionContext);
            var roleText = this._roleText.Get(executionContext); ;

            try
            {
                tracer.Trace("propertyRef ID: {0}", propertyRef.Id.ToString());
                tracer.Trace("BURef ID: {0}", BURef.Id.ToString());
                tracer.Trace("roleText: {0}", roleText);

                Entity BUEn = service.Retrieve(BURef.LogicalName, BURef.Id, new ColumnSet("name"));
                string BUName = BUEn.GetAttributeValue<string>("name");

                tracer.Trace("BUName: {0}", BUName);

                var resultLogicalName = GetLogicalName(roleText, BUName);

                if (string.IsNullOrEmpty(resultLogicalName)) 
                {
                    tracer.Trace("No Logical Name Found");
                    return;
                }

                Entity propertyEn = service.Retrieve(propertyRef.LogicalName, propertyRef.Id, new ColumnSet(resultLogicalName));
                EntityReference userRef = propertyEn.GetAttributeValue<EntityReference>(resultLogicalName);

                if (userRef == null) 
                {
                    tracer.Trace("User not exist");
                    return;
                }

                this._output.Set(executionContext, new EntityReference(userRef.LogicalName, userRef.Id));



            }
            catch (Exception ex)
            {
                string msg = $"{this.GetType().Name} of Error Message: \r\n \t {ex.Message}";
                tracer.Trace(msg);
                throw new InvalidPluginExecutionException(msg);
            }
        }

        private string GetLogicalName(string roleText, string BuName) 
        {
            switch (roleText) 
            {
                case "Approver":
                    if (BuName.StartsWith("ML")) 
                    {
                        return "lms_lmfreshmarket";
                    }

                    if (BuName.StartsWith("WE"))
                    {
                        return "lms_lmweflare";
                    }

                    if (BuName.StartsWith("RL"))
                    {
                        return "lms_lmretail";
                    }

                    if (BuName.Contains("MM - Minor Letting"))
                    {
                        return "lms_mmmml";
                    }

                    if (BuName.Contains("MM - Sales Kiosk"))
                    {
                        return "lms_lmsk";
                    }

                    if (BuName.Contains("MM - Sales Venue"))
                    {
                        return "lms_lmsv";
                    }

                    if (BuName.Contains("MM - Casual Letting"))
                    {
                        return "lms_lmcl";
                    }

                    if (BuName.Contains("MM - Anciliary Letting"))
                    {
                        return "lms_lmal";
                    }

                    break;
                case "CD Approver":
                    if (BuName.StartsWith("WE"))
                    {
                        return "lms_dealapproverlmwelfare";
                    }

                    if (BuName.StartsWith("RL"))
                    {
                        return "lms_dealapproverlmretail";
                    }
                    break;
                case "GM":
                    if (BuName.StartsWith("RL"))
                    {
                        return "lms_gmretail";
                    }

                    if (BuName.StartsWith("ML"))
                    {
                        return "lms_gmfreshmarket";
                    }

                    if (BuName.StartsWith("WE"))
                    {
                        return "lms_gmwelfare";
                    }

                    if (BuName.Contains("MM"))
                    {
                        return "lms_gmmm";
                    }
                    break;
                case "AGM":
                    if (BuName.StartsWith("ML"))
                    {
                        return "lms_agmfreshmarket";
                    }
                    
                    break;
            }
            return "";
        }
    }
}

