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
    public class LeadPopulateBudgetFields : CodeActivity
    {

        [RequiredArgument]
        [Input("Lease")]
        [ReferenceTarget("lms_lease")]
        public InArgument<EntityReference> _leaseRef { get; set; }

        [RequiredArgument]
        [Input("Lead")]
        [ReferenceTarget("clf_lead")]
        public InArgument<EntityReference> _leadRef { get; set; }
        protected override void Execute(CodeActivityContext executionContext)
        {
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.InitiatingUserId);
            ITracingService tracer = executionContext.GetExtension<ITracingService>();


            try
            {
                EntityReference leaseRef = this._leaseRef.Get(executionContext);
                EntityReference leadRef = this._leadRef.Get(executionContext);
                string budgetYear = "";

                DateTime budgetLeaseStart = new DateTime();
                DateTime budgetLeaseEnd = new DateTime();
                decimal budgetBaseRent = 0;
                decimal budgetMgtFee = 0;


                if (leadRef == null)
                {
                    tracer.Trace("No Lead Found");
                    return;
                }

                if (leaseRef == null)
                {
                    tracer.Trace("No Lease Found");
                    return;
                }
                //Find Budget Year

                QueryExpression systemConfigQe = new QueryExpression("lms_systemconfiguration");
                systemConfigQe.Criteria.AddCondition("lms_name", ConditionOperator.Equal, "Budget_Year");
                systemConfigQe.ColumnSet.AddColumns("lms_value");

                EntityCollection systemConfigEc = service.RetrieveMultiple(systemConfigQe);

                foreach (Entity systemConfigEn in systemConfigEc.Entities)
                {
                    budgetYear = systemConfigEn.GetAttributeValue<string>("lms_value");
                }

                if (string.IsNullOrEmpty(budgetYear)) 
                {
                    tracer.Trace("Budget Year Not Found");
                    return;
                }

                // Find Lead Unit Under Lead

                QueryExpression leadUnitQe = new QueryExpression("lms_leadunit");
                leadUnitQe.Criteria.AddCondition("lms_lead", ConditionOperator.Equal, leadRef.Id);
                leadUnitQe.Criteria.AddCondition("lms_unit", ConditionOperator.NotNull);
                leadUnitQe.ColumnSet.AddColumns("lms_unit");

                EntityCollection leadUnitEc = service.RetrieveMultiple(leadUnitQe);

                if (leadUnitEc.Entities.Count == 0) 
                {
                    tracer.Trace("No Lease Unit Found");
                    return;
                }

                List<EntityReference> unit_list = new List<EntityReference>();

                foreach (Entity leadUnitEn in leadUnitEc.Entities)
                {
                    unit_list.Add(leadUnitEn.GetAttributeValue<EntityReference>("lms_unit"));
                }

                //Find Budgets
                int count = 1;

                foreach (EntityReference unitRef in unit_list) 
                {
                    QueryExpression budgetQe = new QueryExpression("lms_budget");
                    budgetQe.Criteria.AddCondition("lms_unitid", ConditionOperator.Equal, unitRef.Id);
                    budgetQe.Criteria.AddCondition("lms_leaseid", ConditionOperator.Equal, leaseRef.Id);
                    budgetQe.LinkEntities.Add(new LinkEntity("lms_budget", "lms_budgetyear", "lms_year", "lms_budgetyearid", JoinOperator.Inner));
                    budgetQe.LinkEntities[0].LinkCriteria.AddCondition("lms_name", ConditionOperator.Like, budgetYear);

                    budgetQe.ColumnSet.AddColumns("lms_renewaldoc", "lms_renewaldoe", "lms_effectivebudgetunit", "lms_renewalmgtunit");

                    EntityCollection budgetEc = service.RetrieveMultiple(budgetQe);

                    Entity budgetEn = budgetEc.Entities.FirstOrDefault();

                    if (budgetEn != null) 
                    {
                        if (count == 1) 
                        {
                            budgetLeaseStart = budgetEn.GetAttributeValue<DateTime>("lms_renewaldoc");
                            budgetLeaseEnd = budgetEn.GetAttributeValue<DateTime>("lms_renewaldoe");
                        }

                        budgetBaseRent += budgetEn.GetAttributeValue<Money>("lms_effectivebudgetunit") == null?0: budgetEn.GetAttributeValue<Money>("lms_effectivebudgetunit").Value;
                        budgetMgtFee += budgetEn.GetAttributeValue<Money>("lms_renewalmgtunit") == null ? 0 : budgetEn.GetAttributeValue<Money>("lms_renewalmgtunit").Value;

                        //Create Unit Budget

                        Entity new_UnitBudgetEn = new Entity("lms_unitbudget");
                        new_UnitBudgetEn["lms_lead"] = leadRef;
                        new_UnitBudgetEn["lms_budget"] = budgetEn.ToEntityReference();

                        service.Create(new_UnitBudgetEn);

                    }

                    count += 1;
                }

                //Map Fields to Lead
                Entity update_LeadEn = new Entity(leadRef.LogicalName, leadRef.Id);
                update_LeadEn["clf_budgetleasestart"] = budgetLeaseStart;
                update_LeadEn["clf_budgetleaseend"] = budgetLeaseEnd;
                update_LeadEn["clf_budgetbaserent"] = new Money(budgetBaseRent);
                update_LeadEn["clf_budgetmgtfee"] = new Money(budgetMgtFee);

                service.Update(update_LeadEn);

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

