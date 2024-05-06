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
    public class SetWarrantyWithDeleteProducts : CodeActivity
    {
        #region Parameters
        // In this section input and output parameters were defined


        [RequiredArgument]
        [Input("Quote")]
        [ReferenceTarget("quote")]
        public InArgument<EntityReference> quoteField { get; set; }


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
            EntityReference QuoteValue = quoteField.Get<EntityReference>(context);


            using (var crm = new XrmServiceContext(service))
            {
                
                var des = "";
                var quotedetailsDelete = crm.QuoteDetailSet.Where(x => x.Id == workflowContext.PrimaryEntityId).First();
                var Quote = crm.QuoteSet.Where(x => x.Id == QuoteValue.Id).First();
                var QuoteProduct = crm.QuoteDetailSet.Where(x => (x.QuoteId.Id == Quote.Id) && (x.ProductId.Id != quotedetailsDelete.ProductId.Id));
                Quote.rhs_DescriptionDevicewarranty = null;
                Quote.rhs_DescriptionBattarywarranty = null;
                Quote.rhs_DescriptionNonwarrantyBattary = null;
                foreach (var Item in QuoteProduct)
                {
                    var Product = crm.ProductSet.Where(x => x.Id == Item.ProductId.Id).First();
                    if (Item.rhs_WarrantyPeriod != null)
                    {
                        if (Product.rhs_TypeProducts.Name == "یو پی اس" || Product.rhs_TypeProducts.Name == "استابلایزر" || Product.rhs_TypeProducts.Name == "اینورتر" || Product.rhs_TypeProducts.Name == "شارژر" || Product.rhs_TypeProducts.Name == "سوئیچ ATS") //دستگاه
                        {
                            des = "دستگاه " + Item.ProductId.Name + " دارای " + Item.rhs_WarrantyPeriod.Value + " ماه گارانتی، ";
                            Quote.rhs_DescriptionDevicewarranty = Quote.rhs_DescriptionDevicewarranty + des;

                        }
                        else if ((Product.rhs_TypeProducts.Name == "باتری" || Product.rhs_TypeProducts.Name == "پک باتری") && (Product.rhs_ProductSerie.Name != "Saba"))//باتری بغییر از برند صبا
                        {
                            des = "" + Item.ProductId.Name + " دارای " + Item.rhs_WarrantyPeriod.Value + " ماه گارانتی، ";
                            Quote.rhs_DescriptionBattarywarranty = Quote.rhs_DescriptionBattarywarranty + des;
                        }
                        else if ((Product.rhs_TypeProducts.Name == "باتری" || Product.rhs_TypeProducts.Name == "پک باتری") && Product.rhs_ProductSerie.Name == "Saba" && Item.rhs_WarrantyPeriod.Value != 0)//باتری صبا دارای گارانتی
                        {
                            des = "" + Item.ProductId.Name + " دارای " + Item.rhs_WarrantyPeriod.Value + " ماه گارانتی، ";
                            Quote.rhs_DescriptionBattarywarranty = Quote.rhs_DescriptionBattarywarranty + des;
                        }
                        else if ((Product.rhs_TypeProducts.Name == "باتری" || Product.rhs_TypeProducts.Name == "پک باتری") && Product.rhs_ProductSerie.Name == "Saba" && Item.rhs_WarrantyPeriod.Value == 0)//باتری صبا با گارانتی صفر
                        {
                            Quote.rhs_DescriptionNonwarrantyBattary = ".گارانتی و مسئولیت کیفی باتری های برند صبا، صرفا به عهده شرکت سازنده (صبا باتری) می باشد -";
                        }
                    }
                }
                if (Quote.rhs_DescriptionDevicewarranty != null)
                {
                    Quote.rhs_DescriptionDevicewarranty = "." + Quote.rhs_DescriptionDevicewarranty + " و دارای 10 سال خدمات پس از فروش می باشد -";
                }
                if (Quote.rhs_DescriptionBattarywarranty != null)
                {
                    Quote.rhs_DescriptionBattarywarranty = "." + Quote.rhs_DescriptionBattarywarranty + " می باشد -";
                }

                crm.UpdateObject(Quote);
                crm.SaveChanges();   
            }
            #endregion
        }
    }
}

