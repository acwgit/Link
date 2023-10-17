using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace ACW.Plugin
{
    public class CreateAutoNumberFromPrefix : CodeActivity
    {
        [RequiredArgument]
        [Input("Table Name")]
        public InArgument<string> _table { get; set; }

        [RequiredArgument]
        [Input("Prefix")]
        public InArgument<string> _prefix { get; set; }

        [RequiredArgument]
        [Input("Date Format")]
        public InArgument<string> _dateFormat { get; set; }

        [RequiredArgument]
        [Input("Number Format")]
        public InArgument<string> _numberFormat { get; set; }

        [Output("Auto Number")]
        public OutArgument<string> _autoNumber { get; set; }

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
                string table = this._table.Get(executionContext);
                string prefix = this._prefix.Get(executionContext);
                string dateFormat = this._dateFormat.Get(executionContext);
                string numberFormat = this._numberFormat.Get(executionContext);

                tracer.Trace("table: {0}", table);
                tracer.Trace("prefix: {0}", prefix);
                tracer.Trace("dateFormat: {0}", dateFormat);
                tracer.Trace("number: {0}", numberFormat);

                int count = GetTableCount(table);
                string now = convertToLocalTime(userId, DateTime.Now.ToUniversalTime()).ToString(dateFormat);
                string autoNumber = prefix + now + string.Format(numberFormat, count);

                this._autoNumber.Set(executionContext, autoNumber);
            }
            catch (Exception ex)
            {
                string msg = $"{this.GetType().Name} of Error Message: \r\n \t {ex.Message}";
                tracer.Trace(msg);
                throw new InvalidPluginExecutionException(msg);
            }
        }

        private int GetTableCount(string table)
        {
            QueryExpression qe = new QueryExpression(table);
            qe.Criteria.AddCondition("createdon", ConditionOperator.Today);
            EntityCollection ec = service.RetrieveMultiple(qe);
            tracer.Trace("Table Count: {0}", ec.Entities.Count.ToString());
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
