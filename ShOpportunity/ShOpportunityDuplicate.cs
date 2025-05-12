using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using ShPlugins;
using System;
using System.Linq;
using System.ServiceModel;

namespace ShOpportunity
{
    public class ShOpportunityDuplicate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Initiate "context", "service" and "tracing".
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                // Trace log to verify plugin is running (before context variable).
                tracing.Trace("plugin> Verify plugin is running: ShOpportunityDuplicate");
                // Check for "Target" variable and check for "Target" type is "EntityReference" before plugin begin to run.
                // This uses "Target" because the plugin is linked to a Bound Custom Action, "Target" is generated automatically.
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference)
                {
                    // This plugin is activated using Bound Custom Action, so "EntityReference" is used instead of "Entity".
                    EntityReference entityRef = (EntityReference)context.InputParameters["Target"];
                    if (entityRef.LogicalName != Opportunity.EntityLogicalName) return;
                    // The "EntityReference" can only gives the logical name and the GUID.
                    string entityRefLogicalName = entityRef.LogicalName.ToString();
                    Guid entityRefGUID = entityRef.Id;
                    // Use "RetrieveOriginalOpportunity" to retrieve the full "Opportunity" data based on the "Opportunity" GUID.
                    Entity original = RetrieveOriginalOpportunity(service, tracing, entityRefLogicalName, entityRefGUID);

                    // Cloning process
                    var clone = new Opportunity();
                    var props = typeof(Opportunity).GetProperties();
                    foreach (var prop in props)
                    {
                        if (!prop.CanWrite ||
                            !prop.CanRead ||
                            prop.GetIndexParameters().Length > 0)
                            continue;
                        if ((prop.Name.EndsWith("Id", StringComparison.OrdinalIgnoreCase) &&
                             prop.Name != "CustomerId" &&
                             prop.Name != "ParentContactId" &&
                             prop.Name != "ParentAccountId" &&
                             prop.Name != "OwnerId" &&
                             prop.Name != "TransactionCurrencyId" && prop.Name != "PriceLevelId") ||
                            prop.Name.StartsWith("Created", StringComparison.OrdinalIgnoreCase) ||
                            prop.Name.StartsWith("Modified", StringComparison.OrdinalIgnoreCase) ||
                            prop.Name == "OpportunityId" ||
                            prop.Name == "EntityState" ||
                            prop.Name == "StateCode" ||
                            prop.Name == "StatusCode" ||
                            prop.Name == "Attributes")
                            continue;

                        // Use this for debugging which attribute is duplicate
                        //tracing.Trace("Value: {0}", prop.Name.ToString());

                        var value = prop.GetValue(original);
                        if (value != null)
                        {
                            prop.SetValue(clone, value);
                        }
                    }
                    clone.Name = "[Cloned] " + clone.Name;
                    var clonedId = service.Create(clone);
                    tracing.Trace("Verify cloning Opportunity completed: {0}", clonedId);

                    // 🔽 START CLONE PRODUCTS
                    // Clone the Opportunity Products using FetchXML
                    string fetchProductsXml = $@"
                        <fetch>
                          <entity name='opportunityproduct'>
                            <all-attributes />
                            <filter>
                              <condition attribute='opportunityid' operator='eq' value='{original.Id}' />
                            </filter>
                          </entity>
                        </fetch>";
                    var originalProducts = service.RetrieveMultiple(new FetchExpression(fetchProductsXml));
                    if (originalProducts.Entities.Count == 0)
                    {
                        tracing.Trace("No products found on original opportunity. Skipping product cloning.");
                    }
                    else
                    {

                        tracing.Trace("Cloning {0} opportunity products...", originalProducts.Entities.Count);
                        // Loop through and clone each product
                        foreach (var product in originalProducts.Entities)
                        {
                            Entity clonedProduct = new Entity("opportunityproduct");

                            foreach (var attr in product.Attributes)
                            {
                                // Skip fields that should not be copied
                                if (attr.Key == "opportunityid" ||
                                    attr.Key.EndsWith("id") && attr.Key != "uomid" && attr.Key != "productid")
                                    continue;

                                if (attr.Key.StartsWith("created") ||
                                    attr.Key.StartsWith("modified") ||
                                    attr.Key == "opportunityproductid" ||
                                    attr.Key == "entityimage" ||
                                    attr.Key == "entityimage_timestamp")
                                    continue;

                                clonedProduct[attr.Key] = attr.Value;
                            }
                            // Set the new Opportunity reference
                            clonedProduct["opportunityid"] = new EntityReference("opportunity", clonedId);
                            // Let system recalculate totals; remove calculated fields if needed
                            string[] calcFields = { "baseamount", "extendedamount", "tax", "manualdiscountamount" };
                            foreach (var field in calcFields)
                            {
                                if (clonedProduct.Attributes.Contains(field))
                                    clonedProduct.Attributes.Remove(field);
                            }

                            service.Create(clonedProduct);
                        }
                        tracing.Trace("Finished cloning opportunity products");
                    }
                    // END CLONE PRODUCTS


                    // This process is enough, it has already created Opportunity and BPF with all of the corresponding data
                    // What I need is the stage itself. I can achive this by fetching OpportunitySalesProcess after.

                    // Fetch data from the original BPF
                    string fetchOriginalXml = $@"
                        <fetch top='1'>
                          <entity name='opportunitysalesprocess'>
                            <all-attributes />
                            <filter>
                              <condition attribute='opportunityid' operator='eq' value='{original.Id}' />
                            </filter>
                          </entity>
                        </fetch>";
                    var originalBPF = service.RetrieveMultiple(new FetchExpression(fetchOriginalXml));
                    // Fetch data from the clone BPF
                    string fetchClonedXml = $@"
                        <fetch top='1'>
                          <entity name='opportunitysalesprocess'>
                            <all-attributes />
                            <filter>
                              <condition attribute='opportunityid' operator='eq' value='{clonedId}' />
                            </filter>
                          </entity>
                        </fetch>";
                    var clonedBPF = service.RetrieveMultiple(new FetchExpression(fetchClonedXml));
                    //tracing.Trace(originalBPF.Entities.Count.ToString());
                    //tracing.Trace(clonedBPF.Entities.Count.ToString());
                    // Update the clonedBPF with originalBPF
                    if (originalBPF.Entities.Count > 0)
                    {
                        OpportunitySalesProcess originalBPFType = originalBPF.Entities[0].ToEntity<OpportunitySalesProcess>();
                        OpportunitySalesProcess clonedBPFType = clonedBPF.Entities[0].ToEntity<OpportunitySalesProcess>();
                        var clonedAccount = new OpportunitySalesProcess
                        {
                            Id = clonedBPFType.Id,
                            ProcessId = originalBPFType.ProcessId,
                            //StageId = originalBPFType.StageId, // There isn't any properties called "StageId".
                            TraversedPath = originalBPFType.TraversedPath
                            // Missing the complete option called ActiveStageId and State/Status.
                        };
                        service.Update(clonedAccount);
                    }

                    //// Set the output parameter as "success" once completed
                    ////context.OutputParameters["output"] = "success";

                    // Get the contacts associate with the original account
                    var fetchXml = $@"
                            <fetch>
                              <entity name='crff8_stakeholder'>
                                <link-entity name='crff8_stakeholder_opportunity' from='crff8_stakeholderid' to='crff8_stakeholderid' intersect='true'>
                                  <filter>
                                    <condition attribute='opportunityid' operator='eq' value='{entityRefGUID}' />
                                  </filter>
                                </link-entity>
                              </entity>
                            </fetch>";
                    var contacts = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    //tracing.Trace("Counter: {0}", contacts.Entities.Count.ToString());
                    tracing.Trace("Verify getting contacts");

                    // For each contact add the associate to the clone one
                    foreach (var contact in contacts.Entities)
                    {
                        var associateRequest = new AssociateRequest
                        {
                            Target = new EntityReference("opportunity", clonedId),
                            RelatedEntities = new EntityReferenceCollection
                            {
                                new EntityReference("crff8_stakeholder", contact.Id)
                            },
                            Relationship = new Relationship("crff8_Stakeholder_Opportunity_Opportunity")
                        };
                        service.Execute(associateRequest);

                        //var disassociateRequest = new DisassociateRequest
                        //{
                        //    Target = new EntityReference("crff8_scaccount", entityRefGUID),
                        //    RelatedEntities = new EntityReferenceCollection
                        //    {
                        //        new EntityReference("crff8_sccontact", contact.Id)
                        //    },
                        //    Relationship = new Relationship("crff8_SCContact_crff8_SCAccount_crff8_SCAccount")
                        //};
                        //service.Execute(disassociateRequest);
                    }
                    tracing.Trace("Verify associate contacts with opportunity");

                    // Set the output parameter as "success" once completed
                    //context.OutputParameters["output"] = "success";
                    context.OutputParameters["output"] = clonedId.ToString();
                    tracing.Trace(context.OutputParameters["output"].ToString());
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                // This is for `Retrieve()` is null case
                if (ex.Detail.ErrorCode == -2147220969)
                {
                    tracing.Trace("Exception: the Retrieve() method returns null.");
                    throw new InvalidPluginExecutionException("Exception: the Retrieve() method returns null.");
                }
                // Others
                tracing.Trace("FaultException Code: {0}", ex.Detail.ErrorCode.ToString());
                tracing.Trace("FaultException Message: {0}", ex.Message.ToString());
                throw new InvalidPluginExecutionException("There is FaultException.", ex);
            }
            catch (Exception ex)
            {
                tracing.Trace("plugin> MainExecution - Exception Message: {0}", ex.Message.ToString());
                // Re-throw the (previous) exception, doesn't create and throw a new exception.
                throw; // Give info about which function cause the error without given too much information.
                // All exceptions thrown here only have string value, they don't contain the "ex" value. If you want the "ex" value, make it inline with $"".
                //throw new InvalidPluginExecutionException($"MainExecution - Exception Message: {ex.Message.ToString()}"); // Get most recent exception.
                //throw new InvalidPluginExecutionException($"MainExecution - Exception Message: {ex.InnerException.Message.ToString()}"); // Get the origin exception.
            }
        }

        public Opportunity RetrieveOriginalOpportunity(IOrganizationService service, ITracingService tracing, String entityLogicalName, Guid entityGuid)
        {
            try
            {
                QueryExpression query = new QueryExpression(entityLogicalName)
                {
                    ColumnSet = new ColumnSet(true), // Retrieve all of the column.
                    Criteria = new FilterExpression
                    {
                        Conditions =
                    {
                        new ConditionExpression("opportunityid", ConditionOperator.Equal, entityGuid)
                    }
                    }
                };
                var result = service.RetrieveMultiple(query);
                if (result.Entities.Count == 0)
                {
                    tracing.Trace("plugin> RetrieveOriginalOpportunity - No Opportunity found for Guid: {0}", entityGuid);
                    throw new InvalidPluginExecutionException($"RetrieveOriginalOpportunity - No Opportunity found with Guid: {entityGuid}");
                }
                tracing.Trace("plugin> RetrieveOriginalOpportunity - Opportunity Found for Guid: {0}", entityGuid);
                return result.Entities.FirstOrDefault() as Opportunity;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                tracing.Trace("plugin> RetrieveOriginalOpportunity - Fault Exception Code: {0} - Message: {1}", ex.Detail.ErrorCode, ex.Detail.Message);
                throw new InvalidPluginExecutionException("RetrieveOriginalOpportunity - Fault Exception", ex);
            }
            catch (Exception ex)
            {
                tracing.Trace("plugin> RetrieveOriginalOpportunity - Unexpected Error: {0}", ex.Message.ToString());
                throw new InvalidPluginExecutionException("RetrieveOriginalOpportunity - Unexpected Error", ex);
            }
        }

        public Guid CloneOpportunity(IOrganizationService service, ITracingService tracing, Entity originalOpportunity)
        {
            try
            {
                // ==========================
                // Clone "Opportunity" process.
                // ==========================
                // Prepare variable "clone" to store "Opportunity" values.
                var clone = new Opportunity();
                // Prepare variable "props" for original "Opportunity" properties.
                var props = typeof(Opportunity).GetProperties();
                // Loop through each property of "Opportunity".
                foreach (var prop in props)
                {
                    // Skip these properties because they cause duplication error.
                    if (!prop.CanWrite ||
                        !prop.CanRead ||
                        prop.GetIndexParameters().Length > 0)
                        continue;
                    if ((prop.Name.EndsWith("Id", StringComparison.OrdinalIgnoreCase) &&
                         prop.Name != "CustomerId" &&
                         prop.Name != "ParentContactId" &&
                         prop.Name != "ParentAccountId" &&
                         prop.Name != "OwnerId" &&
                         prop.Name != "TransactionCurrencyId" && prop.Name != "PriceLevelId") ||
                        prop.Name.StartsWith("Created", StringComparison.OrdinalIgnoreCase) ||
                        prop.Name.StartsWith("Modified", StringComparison.OrdinalIgnoreCase) ||
                        prop.Name == "OpportunityId" ||
                        prop.Name == "EntityState" ||
                        prop.Name == "StateCode" ||
                        prop.Name == "StatusCode" ||
                        prop.Name == "Attributes")
                        continue;
                    // Use this to list all of the properties and to debug which properties/attributes cause duplication error.
                    //tracing.Trace("Property Name: {0}", prop.Name.ToString());
                    // Get the value from the property, correspondingly.
                    var value = prop.GetValue(originalOpportunity);
                    // Set the value to the "clone" variable property, correspondingly.
                    if (value != null)
                    {
                        prop.SetValue(clone, value);
                    }
                }
                // Add "[Cloned]" pre-fix to the "clone"'s name property. 
                clone.Name = "[Cloned] " + clone.Name;
                // Create the clone "Opportunity" and assign the GUID value to a variable.
                var clonedId = service.Create(clone);
                tracing.Trace("plugin> CloneOpportunity - Opportunity cloned: {0}", clonedId);



                return new Guid();
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                tracing.Trace("plugin> CloneOpportunity - Fault Exception Code: {0} - Message: {1}", ex.Detail.ErrorCode, ex.Detail.Message);
                throw new InvalidPluginExecutionException("CloneOpportunity - Fault Exception", ex);
            }
            catch (Exception ex)
            {
                tracing.Trace("plugin> CloneOpportunity - Unexpected Error: {0}", ex.Message.ToString());
                throw new InvalidPluginExecutionException("CloneOpportunity - Unexpected Error", ex);
            }
        }

        public void CloneAndAssociateOpportunityProducts(IOrganizationService service, ITracingService tracing)
        {

        }
    }
}
