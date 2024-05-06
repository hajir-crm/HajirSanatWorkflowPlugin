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
    partial class SentSignatureٍExpertImage : CodeActivity
    {
        #region Parameters
        // In this section input and output parameters were defined

        //[RequiredArgument]
        //[Input("Quote")]
        //[ReferenceTarget("quote")]
        //public InArgument<EntityReference> quoteField { get; set; }

        [RequiredArgument]
        [Input("Owner")]
        [ReferenceTarget("systemuser")]
        public InArgument<EntityReference> OwnerField { get; set; }


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
            EntityReference OwnerValue = OwnerField.Get<EntityReference>(context);


            using (var crm = new XrmServiceContext(service))
            {
                if (workflowContext.Depth < 2)
                {
                    var OwnerItems = crm.SystemUserSet.Where(c => c.Id == OwnerValue.Id).First();
                    var QuoteProduct = crm.QuoteSet.Where(x => x.Id == workflowContext.PrimaryEntityId).First();
                   
                    //throw new InvalidWorkflowException("test");
                    var NoteImage = crm.AnnotationSet.Where(c => c.Product_Annotation.Id == OwnerItems.Id).First();


                    string entitytype = "quotedetail";
                    Entity Note = new Entity("annotation");
                    Guid EntityToAttachTo = workflowContext.PrimaryEntityId; // The GUID of the quotedetail
                    Note["objectid"] = new Microsoft.Xrm.Sdk.EntityReference(entitytype, EntityToAttachTo);
                    Note["objecttypecode"] = entitytype;
                    Note["subject"] = NoteImage.Subject;
                    Note["notetext"] = NoteImage.NoteText;
                    Note["documentbody"] = NoteImage.DocumentBody;
                    service.Create(Note);

         
                    
                }
            }
            #endregion
        }
        }
    }
