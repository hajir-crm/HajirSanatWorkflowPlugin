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
    public class GetDatasheetProducts : CodeActivity
    {
        #region Parameters
        // In this section input and output parameters were defined

        //[RequiredArgument]
        //[Input("Quote")]
        //[ReferenceTarget("quote")]
        //public InArgument<EntityReference> quoteField { get; set; }

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

            //Dictionary<int, string> Serise = new Dictionary<int, string>();
            //Serise.Add(130770000, "Homa");
            //Serise.Add(130770001, "Classic I");
            //Serise.Add(130770002, "Classic RMI");
            //Serise.Add(130770003, "Classic");
            //Serise.Add(130770004, "Genesis");
            //Serise.Add(130770005, "Genesis B");
            //Serise.Add(130770006, "Genesis A");
            //Serise.Add(130770007, "Genesis RMI");
            //Serise.Add(130770008, "Genesis RM");
            //Serise.Add(130770009, "Uranus");
            //Serise.Add(130770010, "Eternal");
            //Serise.Add(130770011, "Super Nova");
            //Serise.Add(130770012, "Spider Net");
            //Serise.Add(130770013, "Salicru");
            //Serise.Add(130770014, "Euro Inverter");
            //Serise.Add(130770015, "AVR");
            //Serise.Add(130770016, "STB");
            //Serise.Add(130770017, "STB-3P");
            //Serise.Add(130770018, "Servo");
            //Serise.Add(130770019, "First Power");
            //Serise.Add(130770020, "Piltan");
            //Serise.Add(130770021, "MISOL");
            //Serise.Add(130770022, "Saba");
            //Serise.Add(130770023, "Hajir");
            //Serise.Add(130770024, "Piltan");
            //Serise.Add(130770025, "-");
            //EntityReference QuoteValue = quoteField.Get<EntityReference>(context);
            EntityReference ProductValue = productField.Get<EntityReference>(context);


            using (var crm = new XrmServiceContext(service))
            {
                if (workflowContext.Depth < 2)
                { 
                    var ProductItems = crm.ProductSet.Where(c => c.Id == ProductValue.Id).First();
                    var QuoteProduct = crm.QuoteDetailSet.Where(x => x.Id == workflowContext.PrimaryEntityId).First();

                    //var NoteImage = crm.AnnotationSet.Where(c => c.Product_Annotation.Id == ProductValue.Id).First();
                    //string ProductSerise = Serise[ProductItems.rhs_ProductSeries.Value];
                    //int productserise = ProductItems.rhs_ProductSeries.Value; 
                    string ProductSerise = ProductItems.rhs_ProductSerie.Name;
                    //if (NoteImage != null)
                    //{
                    //    QuoteProduct.rhs_ImageProduct = NoteImage.DocumentBody;
                    //}

                    //throw new InvalidWorkflowException(NoteImage.DocumentBody);

                    if (ProductItems.rhs_TypeProduct.Value == 130770001 )  /*Satblizer*/
                    {
                        QuoteProduct.rhs_ProductDataSheet = ProductItems.Name +
                        "\n" + "Series: " + ProductSerise +
                        "\n" + "Manufatcturer: " + ProductItems.rhs_Manufacturer +
                        "\n" + "Capacity: " + ProductItems.rhs_Capacity +
                        "\n" + "Input Voltage range: " + ProductItems.rhs_InputVoltagerange +
                        "\n" + "Output Voltage: " + ProductItems.rhs_OutputVoltageStablizer +
                        "\n" + "Operating Frequency range: " + ProductItems.rhs_OperatingFrequencyrange +
                        "\n" + "Output Current:" + ProductItems.rhs_OutputCurrent +
                        "\n" + "Dimension WxDxH / mm: " + ProductItems.rhs_DimensionWxDxHmmStablizer +
                        "\n" + "Weight / kg: " + ProductItems.rhs_WeightkgStablizer;
                    }
                    else if (ProductItems.rhs_TypeProduct.Value == 130770000 )
                    {
                        if(QuoteProduct.rhs_TypeDataSheet.Value == 130770001)//Simple
                        {
                            QuoteProduct.rhs_ProductDataSheet = ProductItems.Name +
                           "\n" + "Series: " + ProductSerise +
                           "\n" + "Capacity VA/Watts: " + ProductItems.rhs_CapacityVAWatts +
                           "\n" + "Input Nominal Voltage & Input Operating Voltage range: " + ProductItems.rhs_InputNominalVoltageInputOperatingVoltager +
                           "\n" + "Output Voltage: " + ProductItems.rhs_OutputVoltage +
                           "\n" + "Battery Voltage: " + ProductItems.rhs_BatteryVoltageProtection +
                           "\n" + "Communication Interface: " + ProductItems.rhs_CommunicationInterface +
                           "\n" + "Dimension WxDxH / mm: " + ProductItems.rhs_DimensionWxDxHmm +
                           "\n" + "Weight / kg: " + ProductItems.rhs_Weightkg;
                        }
                        else if(QuoteProduct.rhs_TypeDataSheet.Value == 130770000)//Full
                        {
                            QuoteProduct.rhs_ProductDataSheet = ProductItems.Name +
                           "\n" + "Series: " + ProductSerise +
                           "\n" + "Manufatcturer: " + ProductItems.rhs_Manufacturer +
                           "\n" + "Capacity VA/Watts: " + ProductItems.rhs_CapacityVAWatts +
                           "\n" + "Technology: " + ProductItems.rhs_Technology +
                           "\n" + "Input Nominal Voltage & Input Operating Voltage range: " + ProductItems.rhs_InputNominalVoltageInputOperatingVoltager +
                           "\n" + "Input Operating Frequency range: " + ProductItems.rhs_InputOperatingFrequencyrange +
                           "\n" + "Input Power Factor: " + ProductItems.rhs_InputPowerFactor +
                           "\n" + "Output Voltage: " + ProductItems.rhs_OutputVoltage +
                           "\n" + "Output Frequency: " + ProductItems.rhs_OutputFrequency +
                           "\n" + "Output Crest Factor: " + ProductItems.rhs_OutputCrestFactor +
                           "\n" + "Output Efficiency: " + ProductItems.rhs_OutputEfficiency +
                           "\n" + "Output Power Factor: " + ProductItems.rhs_OutputPowerFactor +
                           "\n" + "Battery Voltage: " + ProductItems.rhs_BatteryVoltageProtection +
                           "\n" + "Protection: " + ProductItems.rhs_Protection +
                           "\n" + "Communication Interface: " + ProductItems.rhs_CommunicationInterface +
                           "\n" + "Dimension WxDxH / mm: " + ProductItems.rhs_DimensionWxDxHmm +
                           "\n" + "Weight / kg: " + ProductItems.rhs_Weightkg;
                        }
                        
                    }
                    else if (ProductItems.rhs_TypeProduct.Value == 130770002 || ProductItems.rhs_TypeProduct.Value == 130770005 ) /*Battery and Battery Pack*/
                    {
                        QuoteProduct.rhs_ProductDataSheet = ProductItems.Name +
                        "\n" + "Series: " + ProductSerise +
                        "\n" + "Manufatcturer: " + ProductItems.rhs_Manufacturer +
                        "\n" + "Voltage / V: " + ProductItems.rhs_VoltageV +
                        "\n" + "Current / A: " + ProductItems.rhs_CurrentA +
                        "\n" + "Dimension WxDxH / mm: " + ProductItems.rhs_DimensionWxDxHmmBattery +
                        "\n" + "Weight / kg: " + ProductItems.rhs_WeightkgBattery;
                    }
                    else if (ProductItems.rhs_TypeProduct.Value == 130770003) /*Cabinet*/
                    {
                        QuoteProduct.rhs_ProductDataSheet = ProductItems.Name +
                        "\n" + "Series: " + ProductSerise +
                        "\n" + "Manufatcturer: " + ProductItems.rhs_Manufacturer +
                        "\n" + "Number of floors: " + ProductItems.rhs_Numberoffloors +
                        "\n" + "Internal useful dimension WxDxH / mm: " + ProductItems.rhs_InternalusefuldimensionWxDxHmm +
                        "\n" + "External dimensions WxDxH / mm: " + ProductItems.rhs_ExternaldimensionsWxDxHmm;
                    }
                    else if (ProductItems.rhs_TypeProduct.Value == 130770004) /*Inverter*/
                    {
                        if (QuoteProduct.rhs_TypeDataSheet.Value == 130770001)//Simple
                        {
                            QuoteProduct.rhs_ProductDataSheet = ProductItems.Name +
                            "\n" + "Series: " + ProductSerise +
                            "\n" + "Capacity VA/Watts: " + ProductItems.rhs_CapacityVAWattsInverter +
                            "\n" + "Input Nominal Voltage & Input Operating Voltage range: " + ProductItems.rhs_InputNominalVoltageInputOperatingVoltageI +
                            "\n" + "Output Voltage: " + ProductItems.rhs_OutputVoltageInverter +
                            "\n" + "Battery Voltage: " + ProductItems.rhs_BatteryVoltageProtectionInverter +
                            "\n" + "Communication Interface: " + ProductItems.rhs_CommunicationInterfaceInverter +
                            "\n" + "Dimension WxDxH / mm: " + ProductItems.rhs_DimensionWxDxHmmInverter +
                            "\n" + "Weight / kg: " + ProductItems.rhs_WeightkgInverter;
                        }
                        else if (QuoteProduct.rhs_TypeDataSheet.Value == 130770000)//Full
                        {
                            QuoteProduct.rhs_ProductDataSheet = ProductItems.Name +
                            "\n" + "Series: " + ProductSerise +
                            "\n" + "Manufatcturer: " + ProductItems.rhs_Manufacturer +
                            "\n" + "Capacity VA/Watts: " + ProductItems.rhs_CapacityVAWattsInverter +
                            "\n" + "Technology: " + ProductItems.rhs_TechnologyInverter +
                            "\n" + "Input Nominal Voltage & Input Operating Voltage range: " + ProductItems.rhs_InputNominalVoltageInputOperatingVoltageI +
                            "\n" + "Input Operating Frequency range: " + ProductItems.rhs_InputOperatingFrequencyrangeInverter +
                            "\n" + "Input Power Factor: " + ProductItems.rhs_InputPowerFactorInverter +
                            "\n" + "Output Voltage: " + ProductItems.rhs_OutputVoltageInverter +
                            "\n" + "Output Frequency: " + ProductItems.rhs_OutputFrequencyInverter +
                            "\n" + "Output Crest Factor: " + ProductItems.rhs_OutputCrestFactorInverter +
                            "\n" + "Output Efficiency: " + ProductItems.rhs_OutputEfficiencyInverter +
                            "\n" + "Output Power Factor: " + ProductItems.rhs_OutputPowerFactorInverter +
                            "\n" + "Battery Voltage: " + ProductItems.rhs_BatteryVoltageProtectionInverter +
                            "\n" + "Protection: " + ProductItems.rhs_ProtectionInverter +
                            "\n" + "Communication Interface: " + ProductItems.rhs_CommunicationInterfaceInverter +
                            "\n" + "Dimension WxDxH mm: " + ProductItems.rhs_DimensionWxDxHmmInverter +
                            "\n" + "Weight kg: " + ProductItems.rhs_WeightkgInverter;
                        }
                    }
                    else if (ProductItems.rhs_TypeProduct.Value == 130770006 ) /*Switch ATS*/
                    {
                        QuoteProduct.rhs_ProductDataSheet = ProductItems.Name +
                        "\n" + "Series: " + ProductSerise +
                        "\n" + "Manufatcturer: " + ProductItems.rhs_Manufacturer +
                        "\n" + "Input Voltage & Input Operating Voltage range: " + ProductItems.rhs_InputVoltageInputOperatingVoltagerangeATS +
                        "\n" + "Input Operating Frequency range: " + ProductItems.rhs_InputOperatingFrequencyrangeATS +
                        "\n" + "Input Current Max: " + ProductItems.rhs_InputCurrentMaxATS +
                        "\n" + "Output Voltage: " + ProductItems.rhs_OutputVoltageATS +
                        "\n" + "Output Current Max: " + ProductItems.rhs_OutputCurrentMaxATS +
                        "\n" + "Dimension WxDxH mm: " + ProductItems.rhs_DimensionWxDxHmmATS +
                        "\n" + "Weight kg: " + ProductItems.rhs_WeightkgATS;
                    }
                    crm.UpdateObject(QuoteProduct);
                    crm.SaveChanges();
                    

                }
            }
            #endregion
        }
        }
    }
