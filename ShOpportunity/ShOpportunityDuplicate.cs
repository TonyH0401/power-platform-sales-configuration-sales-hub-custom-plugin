using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using ShPlugins;
using System;
using System.ServiceModel;

namespace ShOpportunity
{
    public class ShOpportunityDuplicate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                // Verify the plugin is running
                tracing.Trace("Opportunity: Verify plugin runs");
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference)
                {
                    // Plugin is activated using Custom Action, so 'EntityReference' is used
                    EntityReference entityRef = (EntityReference)context.InputParameters["Target"];
                    if (entityRef.LogicalName != Opportunity.EntityLogicalName) return;
                    // We are using 'EntityReference' so we can get only logical name and GUID
                    string entityRefLogicalName = entityRef.LogicalName.ToString();
                    Guid entityRefGUID = entityRef.Id;
                    // Retrieve the full data
                    Entity original = service.Retrieve(entityRefLogicalName, entityRefGUID, new ColumnSet(true));

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
                             prop.Name != "OwnerId") ||
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
                            //StageId = originalBPFType.StageId,
                            TraversedPath = originalBPFType.TraversedPath
                            // Missing the complete option called Active Stage
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
                tracing.Trace("Exception Message: {0}", ex.Message.ToString());
                throw;
            }
        }
    }
}
