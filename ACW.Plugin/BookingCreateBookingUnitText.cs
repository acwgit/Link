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
    public class BookingCreateBookingUnitText : CodeActivity
    {


        [RequiredArgument]
        [ReferenceTarget("lms_mmbooking")]
        [Input("Booking")]
        public InArgument<EntityReference> _bookingRef { get; set; }

        [RequiredArgument]        
        [Input("Seperator")]
        public InArgument<string> _seperator { get; set; }


        [Output("TextValue")]
        public OutArgument<string> _output { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.InitiatingUserId);
            ITracingService tracer = executionContext.GetExtension<ITracingService>();


            var bookingRef = this._bookingRef.Get(executionContext);
            var seperator = this._seperator.Get(executionContext);

            try
            {
                if (bookingRef == null) 
                {
                    tracer.Trace("bookingRef not found");
                }

                QueryExpression bookingUnitQe = new QueryExpression("lms_bookingunit");
                bookingUnitQe.Criteria.AddCondition("lms_booking", ConditionOperator.Equal, bookingRef.Id);
                bookingUnitQe.Criteria.AddCondition("lms_unit", ConditionOperator.NotNull);
                bookingUnitQe.ColumnSet.AddColumns("lms_unit");

                EntityCollection bookingUnitEc = service.RetrieveMultiple(bookingUnitQe);

                string tempString = "";
                int count = 0;
                
                foreach (Entity bookingUnitEn in bookingUnitEc.Entities) 
                {
                    EntityReference unitRef = bookingUnitEn.GetAttributeValue<EntityReference>("lms_unit");
                    Entity unitEn = service.Retrieve(unitRef.LogicalName, unitRef.Id, new ColumnSet("clf_unitcode"));

                    string unitCode = unitEn.GetAttributeValue<string>("clf_unitcode");

                    if (!string.IsNullOrEmpty(unitCode)) 
                    {
                        if (count == 0)
                        {
                            tempString += unitCode;
                        }
                        else 
                        {
                            tempString += seperator + unitCode;
                        }
                    }

                    count += 1;
                }

                tracer.Trace("final text: {0}", tempString);
                this._output.Set(executionContext, tempString);
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

