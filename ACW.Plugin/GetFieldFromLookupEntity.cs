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
    public class GetFieldFromLookupEntity : CodeActivity
    {
        [RequiredArgument]
        [Input("Record URL")]
        
        public InArgument<string> _recordURL { get; set; }
        [RequiredArgument]
        [Input("Entity logical name")]
        public InArgument<string> _entitylogicalname { get; set; }

        [RequiredArgument]
        [Input("Lookup Entity Field Logical Name")]
        public InArgument<string> _lookupentityFiledLogicalname { get; set; }

        [RequiredArgument]
        [Input("Field logical name")]
        public InArgument<string> _fieldlogicalname { get; set; }

        [RequiredArgument]
        [Input("Data Type (Text,Option,Two Option)")]
        public InArgument<string> _datatype { get; set; }

      
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
            //var _lookupentitylogicalname = this._lookupentitylogicalname.Get(executionContext);
            var _lookupentityFiledLogicalname = this._lookupentityFiledLogicalname.Get(executionContext);


            try
            {
                string RecordURL = this._recordURL.Get(executionContext);
                string[] urlParts = RecordURL.Split("?".ToArray());
                string[] urlParams = urlParts[1].Split("&".ToCharArray());
                string recordId = string.Empty;
                //entityName = urlParams[2].Replace("etn=", "");
                //objectId = urlParams[3].Replace("id=", "");
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

                    if (entity.Attributes.Contains(_lookupentityFiledLogicalname))
                    {
                        EntityReference LookupRecord = (EntityReference)entity.Attributes[_lookupentityFiledLogicalname];
                        if (LookupRecord != null)
                        {
                            //err += "LookupRecord retrieved";
                            Entity lookUpEntity= service.Retrieve(LookupRecord.LogicalName,LookupRecord.Id,new ColumnSet(true));
                            switch (datatype)
                            {
                                case "Text":
                                    if (lookUpEntity.Attributes.Contains(fieldlogicalname))
                                    {
                                        err += "LookupRecord retrieved";
                                        string textValue = lookUpEntity.Attributes[fieldlogicalname].ToString();
                                        this._output.Set(executionContext, textValue);
                                    }
                                    break;
                                case "Option":
                                    
                                    if (lookUpEntity.FormattedValues.Contains(fieldlogicalname))
                                    {

                                        string textValue = lookUpEntity.FormattedValues[fieldlogicalname];
                                        this._output.Set(executionContext, textValue);
                                    }
                                    break;
                                case "Two Option":

                                    if (lookUpEntity.Attributes.Contains(fieldlogicalname))
                                    {
                                        string textValue=string.Empty; 
                                        bool boolValue =(bool)lookUpEntity.Attributes[fieldlogicalname];
                                        if (boolValue)
                                            textValue = "Yes";
                                        else
                                            textValue = "No";
                                        this._output.Set(executionContext, textValue);
                                    }
                                    break;
                                default:
                                    break;
                            }
                            
                           
                            //_output.Set(executionContext, userLookup);
                        }
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

