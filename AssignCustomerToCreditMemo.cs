namespace CompanyA.AssignCustomerToCreditMemo.Workflow
{
    using System;
    using System.Activities;
    using System.ServiceModel;
    using System.Collections.Generic;

    //CRM
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Workflow;
    using Microsoft.Xrm.Sdk.Query;
    using System.Linq;
    using Microsoft.Crm.Sdk.Messages;

    public sealed class AssignCustomerToCreditMemo : CodeActivity
    {
        #region Workflow Parameters       

        [RequiredArgument]
        [Input("Credit Memo")]
        [ReferenceTarget("hdp_creditmemo")]
        public InArgument<EntityReference> CreditMemoReference { get; set; }

        [Output("Customer")]
        [ReferenceTarget("account")]
        public OutArgument<EntityReference> AccountReference { get; set; }

        [Output("EmailToBeSent")]
        public OutArgument<string> EmailToBeSent { get; set; }

        #endregion

        protected override void Execute(CodeActivityContext executionContext)
        {
            #region Local variables

            ITracingService tracingService = null;
            IWorkflowContext context = null;
            IOrganizationServiceFactory serviceFactory = null;
            IOrganizationService service = null;

            string customerNumber = string.Empty;
            bool hasAccount = false;
            Guid accountId = Guid.Empty;
            EntityReference creditMemo = null;

            #endregion

            try
            {
                #region Load Services

                // Create the tracing service
                tracingService = executionContext.GetExtension<ITracingService>();

                if (tracingService == null)
                {
                    throw new InvalidPluginExecutionException("Failed to retrieve tracing service.");
                }

                tracingService.Trace("Entered AssignCustomerToCreditMemo.Execute(), Activity Instance Id: {0}, Workflow Instance Id: {1}",
                    executionContext.ActivityInstanceId,
                    executionContext.WorkflowInstanceId);

                // Create the context
                context = executionContext.GetExtension<IWorkflowContext>();

                if (context == null)
                {
                    throw new InvalidPluginExecutionException("Failed to retrieve workflow context.");
                }

                tracingService.Trace("AssignCustomerToCreditMemo.Execute(), Correlation Id: {0}, Initiating User: {1}",
                    context.CorrelationId,
                    context.InitiatingUserId);

                serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
                service = serviceFactory.CreateOrganizationService(context.UserId);

                #endregion

                creditMemo = CreditMemoReference.Get(executionContext);

                ////Get Customer Number
                Entity creditMemoTarget = service.Retrieve("hdp_creditmemo", creditMemo.Id, new ColumnSet("hdp_customernumber"));
                if (creditMemoTarget.Contains("hdp_customernumber") && creditMemoTarget["hdp_customernumber"] != null)
                {
                    customerNumber = (string)creditMemoTarget["hdp_customernumber"];
                }

                if (customerNumber != null)
                {
                    //Check for account by Serial Number
                    EntityCollection accountEntities = GetaccountByCustomerNumber(customerNumber, service);

                    foreach (var accountEntity in accountEntities.Entities)
                    {
                        accountId = accountEntity.Id;
                        hasAccount = true;
                    }

                    if (hasAccount && accountId != Guid.Empty)
                    {
                        AccountReference.Set(executionContext, new EntityReference("account", accountId));
                        EmailToBeSent.Set(executionContext, "Found Account");
                    }
                }

                if (!hasAccount)
                {
                    EmailToBeSent.Set(executionContext, "Account not found");
                }
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                tracingService.Trace("Exception: {0}", e.ToString());

                throw new Exception("Exception: " + e.ToString());
            }

            tracingService.Trace("Exiting .Execute(), Correlation Id: {0}", context.CorrelationId);
        }

        private EntityCollection GetaccountByCustomerNumber(string accountNumber, IOrganizationService service)
        {
            string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='account'>    
                                    <attribute name='accountid' />
                                    <attribute name='accountnumber' />
                                    <attribute name='statecode' />   
                                    <order attribute='accountnumber' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='accountnumber' operator='eq' value='{0}' />
                                      <condition attribute='statecode' operator='eq' value='0' />
                                    </filter>
                                  </entity>
                                </fetch>
                                ";

            fetchXML = fetchXML.Replace("{0}", accountNumber.ToString());

            EntityCollection entityaccounts = service.RetrieveMultiple(new FetchExpression(fetchXML));

            return entityaccounts;
        }
    }
}
