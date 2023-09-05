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
    public class TaskGetBatchAndAmountText : CodeActivity
    {
        

        [RequiredArgument]
        [ReferenceTarget("task")]
        [Input("Task")]
        public InArgument<EntityReference> _taskRef { get; set; }      


        [Output("TextValue")]
        public OutArgument<string> _output { get; set; }
        
        protected override void Execute(CodeActivityContext executionContext)
        {
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.InitiatingUserId);
            ITracingService tracer = executionContext.GetExtension<ITracingService>();


            var taskRef = this._taskRef.Get(executionContext);           

            try
            {
                Entity taskEn = service.Retrieve(taskRef.LogicalName, taskRef.Id, new ColumnSet("lms_batchno", "lms_batchno2", "lms_batchno3", "lms_batchno4", "lms_batchno5", "lms_batchno6", "lms_batchno7", "lms_batchno8", "lms_amount1", "lms_amount2", "lms_amount3", "lms_amount4", "lms_amount5", "lms_amount6", "lms_amount7", "lms_amount8"));

                string text_output = "";

                for (int i = 0; i < 8; i++) 
                {
                    
                    if (i == 0)
                    {
                        text_output += taskEn.GetAttributeValue<string>("lms_batchno") == null ? "" : $"#{taskEn.GetAttributeValue<string>("lms_batchno")}";

                    }
                    else 
                    {
                        text_output += taskEn.GetAttributeValue<string>("lms_batchno" + (i + 1).ToString()) == null ? "" : $"# {taskEn.GetAttributeValue<string>("lms_batchno" + (i + 1).ToString())}";
                    }

                    text_output += taskEn.GetAttributeValue<Money>("lms_amount" + (i+1).ToString()) == null ? "" : $" (${taskEn.GetAttributeValue<Money>("lms_amount" + (i + 1).ToString()).Value.ToString("0.00")})";

                    if (i == 0)
                    {
                        text_output += taskEn.GetAttributeValue<string>("lms_batchno") == null && taskEn.GetAttributeValue<Money>("lms_amount" + (i + 1).ToString()) == null ? "" : $"\n";
                    }
                    else 
                    {
                        text_output += taskEn.GetAttributeValue<string>("lms_batchno" + (i + 1).ToString()) == null && taskEn.GetAttributeValue<Money>("lms_amount" + (i + 1).ToString()) == null ? "" : $"\n";
                    }

                    
                    tracer.Trace("amount {0}", (i + 1).ToString());
                }

                tracer.Trace("output: {0}", text_output);

                this._output.Set(executionContext, text_output);
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

