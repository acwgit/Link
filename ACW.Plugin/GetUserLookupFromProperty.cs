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
    public class GetUserLookupFromProperty : CodeActivity
    {
        [RequiredArgument]
        [Input("Property Lookup")]
        [ReferenceTarget("clf_property")]
        public InArgument<EntityReference> _propertyRecord { get; set; }
        //[RequiredArgument]
        //[Input("Entity logical name")]
        //public InArgument<string> _entitylogicalname { get; set; }
        [RequiredArgument]
        [Input("User field logical name")]
        public InArgument<string> _userfieldlogicalname { get; set; }

        [ReferenceTarget("systemuser")]
        [Output("User")]
        public OutArgument<EntityReference> _output { get; set; }
        //string err = string.Empty;
        protected override void Execute(CodeActivityContext executionContext)
        {
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.InitiatingUserId);
            ITracingService tracer = executionContext.GetExtension<ITracingService>();

           
            var userlogicalname = this._userfieldlogicalname.Get(executionContext);

          
            try
            {
                EntityReference propertyRecord = this._propertyRecord.Get(executionContext);
               // string entityName = this._entitylogicalname.Get(executionContext);
                //string entityName = "";
                //string objectId = "";
                
                if (propertyRecord != null)
                {
                  
                    Entity entity = service.Retrieve(propertyRecord.LogicalName, propertyRecord.Id, new ColumnSet(true));
                    if (entity.Attributes.Contains(userlogicalname))
                    {
                        EntityReference userLookup = (EntityReference)entity.Attributes[userlogicalname];
                        if (userLookup != null)
                        {
                            this._output.Set(executionContext, new EntityReference(userLookup.LogicalName, userLookup.Id));
                            //_output.Set(executionContext, userLookup);
                        }
                    }

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

