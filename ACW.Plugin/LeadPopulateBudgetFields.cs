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
        [Input("Lead")]
        [ReferenceTarget("clf_lead")]
        public InArgument<EntityReference> _leadRef { get; set; }

        [RequiredArgument]
        [Input("Lease End + 1")]        
        public InArgument<DateTime> _leaseEndPlus1 { get; set; }
        protected override void Execute(CodeActivityContext executionContext)
        {
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.InitiatingUserId);
            ITracingService tracer = executionContext.GetExtension<ITracingService>();


            try
            {
                
                EntityReference leadRef = this._leadRef.Get(executionContext);
                DateTime leaseEndPlus1 = this._leaseEndPlus1.Get(executionContext);
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
                    budgetQe.Criteria.AddCondition("lms_renewaldoc", ConditionOperator.OnOrBefore, leaseEndPlus1);
                    budgetQe.Criteria.AddCondition("lms_renewaldoe", ConditionOperator.OnOrAfter, leaseEndPlus1);
                    budgetQe.LinkEntities.Add(new LinkEntity("lms_budget", "lms_budgetyear", "lms_year", "lms_budgetyearid", JoinOperator.Inner));
                    budgetQe.LinkEntities[0].LinkCriteria.AddCondition("lms_name", ConditionOperator.Like, budgetYear);

                    budgetQe.ColumnSet.AddColumns("lms_renewaldoc", "lms_renewaldoe", "lms_effectivebudgetunit", "lms_renewalmgtunit");

                    EntityCollection budgetEc = service.RetrieveMultiple(budgetQe);

                    Entity budgetEn = budgetEc.Entities.FirstOrDefault();

                    if(budgetEn == null) 
                    {
                        QueryExpression budgetQe2 = new QueryExpression("lms_budget");
                        budgetQe2.Criteria.AddCondition("lms_unitid", ConditionOperator.Equal, unitRef.Id);
                        budgetQe2.Criteria.AddCondition("lms_renewaldoc", ConditionOperator.OnOrAfter, leaseEndPlus1);                        
                        budgetQe2.LinkEntities.Add(new LinkEntity("lms_budget", "lms_budgetyear", "lms_year", "lms_budgetyearid", JoinOperator.Inner));
                        budgetQe2.LinkEntities[0].LinkCriteria.AddCondition("lms_name", ConditionOperator.Like, budgetYear);

                        budgetQe2.AddOrder("lms_renewaldoc", OrderType.Ascending);

                        budgetQe2.ColumnSet.AddColumns("lms_renewaldoc", "lms_renewaldoe", "lms_effectivebudgetunit", "lms_renewalmgtunit");

                        EntityCollection budgetEc2 = service.RetrieveMultiple(budgetQe2);

                        budgetEn = budgetEc2.Entities.FirstOrDefault();
                    }

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
                if (!budgetLeaseStart.ToString().Contains("0001")) 
                {
                    update_LeadEn["clf_budgetleasestart"] = budgetLeaseStart;
                }

                if (!budgetLeaseEnd.ToString().Contains("0001"))
                {
                    update_LeadEn["clf_budgetleaseend"] = budgetLeaseEnd;
                }
               
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

