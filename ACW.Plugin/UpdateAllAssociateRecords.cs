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
    public class UpdateAllAssociateRecords : CodeActivity
    {
        [RequiredArgument]
        [Input("Record URL")]
        
        public InArgument<string> _recordURL { get; set; }
        [RequiredArgument]
        [Input("Entity logical name")]
        public InArgument<string> _entitylogicalname { get; set; }


        [RequiredArgument]
        [Input("Associated Entity Logical Name")]
        public InArgument<string> _associatedEntityLogicalName { get; set; }

        [RequiredArgument]
        [Input("Lookup Field Logical Name in Related Record")]
        public InArgument<string> _associatedLookupFieldLogicalName { get; set; }

        [RequiredArgument]
        [Input("Update Field logical name")]
        public InArgument<string> _fieldlogicalname { get; set; }

        [RequiredArgument]
        [Input("Data Type (Text,Option,Two Option)")]
        public InArgument<string> _datatype { get; set; }

        [RequiredArgument]
        [Input("Value")]
        public InArgument<string> _value { get; set; }


        [Output("TextValue")]
        public OutArgument<string> _output { get; set; }
        string err = string.Empty;
        protected override void Execute(CodeActivityContext executionContext)
        {
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.InitiatingUserId);
            ITracingService tracer = executionContext.GetExtension<ITracingService>();

           
            var fieldlogicalname = this._fieldlogicalname.Get(executionContext);
            var datatype = this._datatype.Get(executionContext);
            var value = this._value.Get(executionContext);
            //var _lookupentitylogicalname = this._lookupentitylogicalname.Get(executionContext);
            var associatedLookupFieldLogicalName = this._associatedLookupFieldLogicalName.Get(executionContext);
            string associatedEntityLogicalName = this._associatedEntityLogicalName.Get(executionContext);
            //string associatedEntityLogicalName = this._associatedEntityLogicalName.Get(executionContext);

            try
            {
                string RecordURL = this._recordURL.Get(executionContext);
                string[] urlParts = RecordURL.Split("?".ToArray());
                string[] urlParams = urlParts[1].Split("&".ToCharArray());
                string recordId = string.Empty;

                for (int i = 0; i < urlParams.Length; i++)
                {

                    if (urlParams[i].Contains("id"))
                    {
                        string[] paramParts = urlParams[i].Split("=".ToArray());
                        recordId = paramParts[1];
                    }

                }
               
                err += recordId;
                string entityName = this._entitylogicalname.Get(executionContext);
                Entity entity=service.Retrieve(entityName,new Guid(recordId),new ColumnSet(true));
                
                if (entity != null)
                {
                    err += "entity retrieved";
                   
                    QueryExpression qeAccount = new QueryExpression(associatedEntityLogicalName);
                    qeAccount.ColumnSet.AddColumns(fieldlogicalname);
                    FilterExpression filter = new FilterExpression(LogicalOperator.And);
                    filter.AddCondition(associatedLookupFieldLogicalName, ConditionOperator.Equal, entity.Id);
                    
                    qeAccount.Criteria.AddFilter(filter);
                    EntityCollection entityCollection = service.RetrieveMultiple(qeAccount);

                    foreach (Entity lookUpEntity in entityCollection.Entities)
                    {
                        err += "datatype: " + datatype;
                        switch (datatype)
                        {
                            
                            case "Text":
                                if (lookUpEntity.Attributes.Contains(fieldlogicalname))
                                {
                                    
                                    lookUpEntity.Attributes[fieldlogicalname] = value;
                                    //this._output.Set(executionContext, textValue);
                                }
                                break;
                            case "Option":

                                if (lookUpEntity.Attributes.Contains(fieldlogicalname))
                                {
                                    lookUpEntity.Attributes[fieldlogicalname] = new OptionSetValue(Convert.ToInt32(value));
                                }
                                else
                                    lookUpEntity.Attributes.Add(fieldlogicalname, new OptionSetValue(Convert.ToInt32(value)));
                                break;
                            case "Two Option":

                                if (lookUpEntity.Attributes.Contains(fieldlogicalname))
                                {
                                    if (value.ToLower() == "no" || value.ToLower() == "false")
                                        lookUpEntity.Attributes[fieldlogicalname] = false;
                                    if (value.ToLower() == "yes" || value.ToLower() == "true")
                                        lookUpEntity.Attributes[fieldlogicalname] = true;

                                }
                                break;
                            default:
                                break;
                        }
                        service.Update(lookUpEntity);
                    }


                }
            }
            catch (Exception ex)
            {
                string msg = $"{this.GetType().Name} of Error Message: \r\n \t {ex.Message}";
                tracer.Trace(msg);
                throw new InvalidPluginExecutionException(msg+ err);
            }
        }
    }
}

