using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace ACW.Plugin
{
    public class BookingPrefixLogic : CodeActivity
    {
        [RequiredArgument]
        [Input("Booking")]
        [ReferenceTarget("lms_mmbooking")]
        public InArgument<EntityReference> _booking { get; set; }

        IOrganizationService service;
        ITracingService tracer;
        Guid userId;

        protected override void Execute(CodeActivityContext executionContext)
        {
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            this.service = serviceFactory.CreateOrganizationService(context.InitiatingUserId);
            this.tracer = executionContext.GetExtension<ITracingService>();
            this.userId = context.InitiatingUserId;

            try
            {
                Entity booking = service.Retrieve("lms_mmbooking", this._booking.Get(executionContext).Id, new ColumnSet("lms_bookingrecordno", "lms_unittype"));
                Entity unitType = service.Retrieve("lms_unittype", booking.GetAttributeValue<EntityReference>("lms_unittype").Id, new ColumnSet("lms_unittypecode"));
                tracer.Trace("booking: {0}", booking.Id.ToString());

                if (unitType == null)
                    return;

                tracer.Trace("unitType: {0}", unitType.Id.ToString());

                string bookingNo = GetSetupPrefix(unitType);
                tracer.Trace("bookingNo: {0}", bookingNo.ToString());
                Entity updateEn = new Entity(booking.LogicalName, booking.Id);
                updateEn["lms_bookingrecordno"] = bookingNo;

                service.Update(updateEn);
            }
            catch (Exception ex)
            {
                string msg = $"{this.GetType().Name} of Error Message: \r\n \t {ex.Message}";
                tracer.Trace(msg);
                throw new InvalidPluginExecutionException(msg);
            }
        }

        private string GetSetupPrefix(Entity unitType)
        {
            string code = unitType.GetAttributeValue<string>("lms_unittypecode").ToUpper();
            string year = convertToLocalTime(userId, DateTime.Now.ToUniversalTime()).ToString("yy");
            int count = GetCount(unitType);
            tracer.Trace("code: {0}", code);
            tracer.Trace("year: {0}", year);
            tracer.Trace("count: {0}", count.ToString());
            string bookingNo = code + year + "-" + string.Format("{0:00000}", count);

            return bookingNo;
        }

        private int GetCount(Entity unitType)
        {
            QueryExpression qe = new QueryExpression("lms_mmbooking");
            qe.Criteria.AddCondition("lms_unittype", ConditionOperator.Equal, unitType.Id);
            qe.Criteria.AddCondition("createdon", ConditionOperator.ThisYear);

            EntityCollection ec = service.RetrieveMultiple(qe);
            return ec.Entities.Count;
        }

        public int RetrieveTimeZoneCodeFromUsersSettings(Guid userid)
        {
            var currentUserSettings = service.RetrieveMultiple(
                new QueryExpression("usersettings")
                {
                    ColumnSet = new ColumnSet("localeid", "timezonecode"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("systemuserid", ConditionOperator.Equal , userid)
                        }
                    }
                }).Entities[0];

            return currentUserSettings.GetAttributeValue<int>("timezonecode");
        }

        public DateTime convertToLocalTime(Guid userid, DateTime utcDate)
        {
            int timeZoneCode = RetrieveTimeZoneCodeFromUsersSettings(userid);
            return convertToLocalTime(timeZoneCode, utcDate);
        }

        public DateTime convertToLocalTime(int timeZoneCode, DateTime utcDate)
        {
            var request = new LocalTimeFromUtcTimeRequest
            {
                TimeZoneCode = timeZoneCode,
                UtcTime = utcDate
            };
            var response2 = service.Execute(request);
            return (DateTime)response2.Results["LocalTime"];
        }
    }
}
