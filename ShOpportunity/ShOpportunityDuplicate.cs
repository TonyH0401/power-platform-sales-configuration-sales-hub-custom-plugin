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
                    // Cloning "Opportunity" process.
                    var clonedId = CloneOpportunity(service, tracing, original);
                    // Set the output parameter value as cloned opportunity GUID for loading form via JS.
                    //context.OutputParameters["output"] = "success";
                    context.OutputParameters["output"] = clonedId.ToString();
                    tracing.Trace("plugin> MainExecution - Opportunity cloning process completed - Cloned opportunity output GUID: {0}", context.OutputParameters["output"].ToString());
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

                // ==========================
                // Clone and associate "Opportunity Product" process. Using late-bound.
                // ==========================
                CloneAndAssociateOpportunityProducts(service, tracing, originalOpportunity, clonedId);

                /*
                 * At this point, this process is enough, a cloned "Opportunity" has already created and due to how the "Opportunity" is bound to the BPF, 
                 * a BPF for the "Opportunity" with all of the corresponding data is also created.
                 * What I need is the stage indication for the BPF itself. I can achieve this by fetching the "OpportunitySalesProcess" data after.
                 */

                // ==========================
                // Update/Clone stage indications from the original opportunity BPF for the cloned opportunity BPF.
                // ==========================
                CloneOpportunityBpfStageIndicators(service, tracing, originalOpportunity, clonedId);

                // ==========================
                // Associate stakeholders with cloned opportunity.
                // ==========================
                AssociateStakeholderOpportunity(service, tracing, originalOpportunity.Id, clonedId);

                // ==========================
                // Verify opportunity clone process is completed and return cloned opportunity id.
                // ==========================
                tracing.Trace("plugin> CloneOpportunity - The opportunity cloning process has been completed: {0}", clonedId);
                return clonedId;
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

        public void CloneAndAssociateOpportunityProducts(IOrganizationService service, ITracingService tracing, Entity originalOpportunity, Guid clonedOpportunityId)
        {
            try
            {
                // Retrieve Opportunity Products using FetchXML.
                string fetchProductsXml = $@"
                        <fetch>
                          <entity name='opportunityproduct'>
                            <all-attributes />
                            <filter>
                              <condition attribute='opportunityid' operator='eq' value='{originalOpportunity.Id}' />
                            </filter>
                          </entity>
                        </fetch>";
                var originalProducts = service.RetrieveMultiple(new FetchExpression(fetchProductsXml));
                // Skip if there are no products. Proceed if there are products.
                if (originalProducts.Entities.Count == 0)
                {
                    tracing.Trace("plugin> CloneAndAssociateOpportunityProducts - No products found on original Opportunity. Skipping product cloning.");
                }
                else
                {

                    tracing.Trace("plugin> CloneAndAssociateOpportunityProducts - {0} opportunity products found. Starting product cloning.", originalProducts.Entities.Count);
                    // Loop through each prodcut and clone them.
                    foreach (var product in originalProducts.Entities)
                    {
                        // Initiate a clone product variable.
                        Entity clonedProduct = new Entity("opportunityproduct");
                        // For each property, clone the value.
                        foreach (var attr in product.Attributes)
                        {
                            // Skip fields that should not be copied.
                            if (attr.Key == "opportunityid" ||
                                attr.Key.EndsWith("id") && attr.Key != "uomid" && attr.Key != "productid")
                                continue;
                            if (attr.Key.StartsWith("created") ||
                                attr.Key.StartsWith("modified") ||
                                attr.Key == "opportunityproductid" ||
                                attr.Key == "entityimage" ||
                                attr.Key == "entityimage_timestamp")
                                continue;
                            // Assign the value of the corresponding property to the clone product.
                            clonedProduct[attr.Key] = attr.Value;
                        }
                        // Set the reference to the newly cloned Opportunity for the newly cloned products. Because product-opportunity has one-many relationship.
                        clonedProduct["opportunityid"] = new EntityReference("opportunity", clonedOpportunityId);
                        // Let the system recalculate totals; remove calculated fields if needed.
                        string[] calcFields = { "baseamount", "extendedamount", "tax", "manualdiscountamount" };
                        foreach (var field in calcFields)
                        {
                            if (clonedProduct.Attributes.Contains(field))
                                clonedProduct.Attributes.Remove(field);
                        }
                        // Create the cloned opportunity products.
                        service.Create(clonedProduct);
                    }
                    tracing.Trace("plugin> CloneAndAssociateOpportunityProducts - Finished cloning and associating opportunity products.");
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                tracing.Trace("plugin> CloneAndAssociateOpportunityProducts - Fault Exception Code: {0} - Message: {1}", ex.Detail.ErrorCode, ex.Detail.Message);
                throw new InvalidPluginExecutionException("CloneAndAssociateOpportunityProducts - Fault Exception", ex);
            }
            catch (Exception ex)
            {
                tracing.Trace("plugin> CloneAndAssociateOpportunityProducts - Unexpected Error: {0}", ex.Message.ToString());
                throw new InvalidPluginExecutionException("CloneAndAssociateOpportunityProducts - Unexpected Error", ex);
            }

        }

        public void CloneOpportunityBpfStageIndicators(IOrganizationService service, ITracingService tracing, Entity originalOpportunity, Guid clonedOpportunityId)
        {
            try
            {
                // Fetch data from the original opportunity BPF with stage indication.
                string fetchOriginalXml = $@"
                <fetch top='1'>
                  <entity name='opportunitysalesprocess'>
                    <all-attributes />
                    <filter>
                      <condition attribute='opportunityid' operator='eq' value='{originalOpportunity.Id}' />
                    </filter>
                  </entity>
                </fetch>";
                var originalBPF = service.RetrieveMultiple(new FetchExpression(fetchOriginalXml));
                // Fetch data from the clone opportunity BPF.
                string fetchClonedXml = $@"
                <fetch top='1'>
                  <entity name='opportunitysalesprocess'>
                    <all-attributes />
                    <filter>
                      <condition attribute='opportunityid' operator='eq' value='{clonedOpportunityId}' />
                    </filter>
                  </entity>
                </fetch>";
                var clonedBPF = service.RetrieveMultiple(new FetchExpression(fetchClonedXml));
                // Update the cloned opportunity BPF with the stage indication from the original opportunity BPF.
                if (originalBPF.Entities.Count > 0)
                {
                    OpportunitySalesProcess originalBPFType = originalBPF.Entities[0].ToEntity<OpportunitySalesProcess>();
                    OpportunitySalesProcess clonedBPFType = clonedBPF.Entities[0].ToEntity<OpportunitySalesProcess>();
                    var clonedStageProcess = new OpportunitySalesProcess
                    {
                        Id = clonedBPFType.Id,
                        ProcessId = originalBPFType.ProcessId,
                        TraversedPath = originalBPFType.TraversedPath
                        //StageId = originalBPFType.StageId, // There isn't any properties called "StageId".
                        // Missing the complete option for BPF called ActiveStageId and State/Status.
                    };
                    // Update the cloned opportunity BPF with the stage indication from the original BPF.
                    service.Update(clonedStageProcess);
                    tracing.Trace("plugin> CloneOpportunityBpfStageIndicators - Finished updating stage indications for cloned opportunity BPF.");
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                tracing.Trace("plugin> CloneOpportunityBpfStageIndicators - Fault Exception Code: {0} - Message: {1}", ex.Detail.ErrorCode, ex.Detail.Message);
                throw new InvalidPluginExecutionException("CloneOpportunityBpfStageIndicators - Fault Exception", ex);
            }
            catch (Exception ex)
            {
                tracing.Trace("plugin> CloneOpportunityBpfStageIndicators - Unexpected Error: {0}", ex.Message.ToString());
                throw new InvalidPluginExecutionException("CloneOpportunityBpfStageIndicators - Unexpected Error", ex);
            }
        }

        public void AssociateStakeholderOpportunity(IOrganizationService service, ITracingService tracing, Guid originalId, Guid clonedId)
        {
            try
            {
                // Get stakeholders associated with the original opportunity via FetchXML.
                var fetchXml = $@"
                    <fetch>
                      <entity name='crff8_stakeholder'>
                        <link-entity name='crff8_stakeholder_opportunity' from='crff8_stakeholderid' to='crff8_stakeholderid' intersect='true'>
                          <filter>
                            <condition attribute='opportunityid' operator='eq' value='{originalId}' />
                          </filter>
                        </link-entity>
                      </entity>
                    </fetch>";
                var stakeholders = service.RetrieveMultiple(new FetchExpression(fetchXml));
                tracing.Trace("plugin> AssociateStakeholderOpportunity - Verify {0} stakeholders retrieved.", stakeholders.Entities.Count.ToString());
                // For each stakeholder, associate it to the cloned opportunity.
                foreach (var stakeholder in stakeholders.Entities)
                {
                    var associateRequest = new AssociateRequest
                    {
                        Target = new EntityReference("opportunity", clonedId),
                        RelatedEntities = new EntityReferenceCollection
                    {
                        new EntityReference("crff8_stakeholder", stakeholder.Id)
                    },
                        Relationship = new Relationship("crff8_Stakeholder_Opportunity_Opportunity")
                    };
                    // Perform the association between the stakeholder and cloned opportunity.
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
                tracing.Trace("plugin> AssociateStakeholderOpportunity - Verify associating stakeholders with cloned opportunity completed.");
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                tracing.Trace("plugin> AssociateStakeholderOpportunity - Fault Exception Code: {0} - Message: {1}", ex.Detail.ErrorCode, ex.Detail.Message);
                throw new InvalidPluginExecutionException("AssociateStakeholderOpportunity - Fault Exception", ex);
            }
            catch (Exception ex)
            {
                tracing.Trace("plugin> AssociateStakeholderOpportunity - Unexpected Error: {0}", ex.Message.ToString());
                throw new InvalidPluginExecutionException("AssociateStakeholderOpportunity - Unexpected Error", ex);
            }
        }
    }
}
