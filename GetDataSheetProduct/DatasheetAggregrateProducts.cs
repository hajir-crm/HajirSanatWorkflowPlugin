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
    public class DatasheetAggregrateProducts : CodeActivity
    {
        #region Parameters
        // In this section input and output parameters were defined

        [RequiredArgument]
        [Input("AggregrateProduct")]
        [ReferenceTarget("rhs_aggregateproducts")]
        public InArgument<EntityReference> AggregrateproductField { get; set; }

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
            EntityReference AggregrateProductValue = AggregrateproductField.Get<EntityReference>(context);


            using (var crm = new XrmServiceContext(service))
            {
                if(AggregrateProductValue != null)//محصول پیش فاکتور ویرایش شده دارای پکیج باشد
                {
                    var des = "";
                    var QuoteItems = crm.QuoteDetailSet.Where(c => c.rhs_Aggregateproducts.Id == AggregrateProductValue.Id);//محصولات وصل شده به پکیج
                    var AggregrateProducts = crm.rhs_aggregateproductsSet.Where(x => x.Id == AggregrateProductValue.Id).First();//پکیج 
                    AggregrateProducts.rhs_ProductDataSheet = null;
                    foreach (var Item in QuoteItems)
                    {
                        var Product = crm.ProductSet.Where(x => x.Id == Item.ProductId.Id).First();
                        if(Item.rhs_ProductDataSheet != null)
                        {
                            des = AggregrateProducts.rhs_ProductDataSheet;
                            AggregrateProducts.rhs_ProductDataSheet = des + Item.rhs_ProductDataSheet + "\n\n\n";
                            //if(Product.rhs_TypeProducts.Name == "یو پی اس" && Item.rhs_ImageProduct != null)
                            //{
                            //    AggregrateProducts.rhs_ImageProduct = Item.rhs_ImageProduct;
                            //}
                        }
                    }
                    crm.UpdateObject(AggregrateProducts);
                    crm.SaveChanges();
                }   
            }
            #endregion
        }
    }
}
