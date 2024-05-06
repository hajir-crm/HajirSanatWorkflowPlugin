using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk.Query;
using System.Activities;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using Xrm;
using System.Reflection;
using System.ComponentModel;
using System.Collections;
using System.Security.Policy;
using System.Web.UI.WebControls;

namespace GetDataSheetProduct
{
    public class CalculateAmount : CodeActivity
    {
        #region Parameters
        // In this section input and output parameters were defined


        [RequiredArgument]
        [Input("Product")]
        [ReferenceTarget("product")]
        public InArgument<EntityReference> productField { get; set; }


        #endregion
        protected override void Execute(CodeActivityContext context)
        {
            IWorkflowContext workflowContext =
              context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory =
              context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service =
              serviceFactory.CreateOrganizationService(workflowContext.UserId);

            #region Manipulation
            // In this section value of the fields from the workflow instance has been get

            //EntityReference QuoteValue = quoteField.Get<EntityReference>(context);
            EntityReference ProductValue = productField.Get<EntityReference>(context);


            using (var crm = new XrmServiceContext(service))
            {
                
                var ProductItems = crm.ProductSet.Where(c => c.Id == ProductValue.Id).First();
                var QuoteProduct = crm.QuoteDetailSet.Where(x => x.Id == workflowContext.PrimaryEntityId).First();

                QuoteProduct.rhs_Amount =Convert.ToDecimal(QuoteProduct.BaseAmount.Value);
                QuoteProduct.rhs_ExtendedAmount = Convert.ToDecimal(QuoteProduct.ExtendedAmount.Value);
                //QuoteProduct.rhs_ManualDiscount = Convert.ToDecimal(QuoteProduct.ManualDiscountAmount.Value);
                QuoteProduct.rhs_PricePerUnit = Convert.ToDecimal(QuoteProduct.PricePerUnit.Value);

                crm.UpdateObject(QuoteProduct);
                crm.SaveChanges();

            }
            #endregion
        }
    }
}
