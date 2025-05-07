using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;

namespace Zox.CreateSalesTeamRecordOnShareUnshare
{
    /// <summary>
    /// Plugin to create or update sales team records when a record (PreLead, Lead, Opportunity) is shared or unshared.
    /// Handles GrantAccess and RevokeAccess messages to manage sales team assignments.
    /// Compatible with .NET Framework 4.6.2 and C# 6.0.
    /// </summary>
    public class CreateSalesTeamRecord : IPlugin
    {
        // Services and context
        private IOrganizationService _service;
        private ITracingService _tracingService;
        private IPluginExecutionContext _context;

        // Entity data
        private Entity _projectEntity;

        // Column sets for retrieving specific attributes
        private readonly ColumnSet _projectColumns = new ColumnSet("zox_project", "zox_package");
        private readonly ColumnSet _opportunityColumns = new ColumnSet("zox_projecthierarchy", "zox_package");
        private readonly ColumnSet _userColumns = new ColumnSet("zox_role", "zox_lob");

        /// <summary>
        /// Main plugin execution method.
        /// </summary>
        /// <param name="serviceProvider">Service provider for accessing CRM services.</param>
        public void Execute(IServiceProvider serviceProvider)
        {
            try
            {
                // Initialize services and context
                InitializeServices(serviceProvider);

                if (_context.MessageName == "GrantAccess")
                {
                    HandleGrantAccess();
                }
                else if (_context.MessageName == "RevokeAccess")
                {
                    HandleRevokeAccess();
                }
            }
            catch (Exception ex)
            {
                _tracingService.Trace("Error in Execute: " + ex.Message);
                throw new InvalidPluginExecutionException("An error occurred in CreateSalesTeamRecord plugin: " + ex.Message, ex);
            }
        }

        /// <summary>
        /// Initializes CRM services and plugin execution context.
        /// </summary>
        private void InitializeServices(IServiceProvider serviceProvider)
        {
            _context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            _service = serviceFactory.CreateOrganizationService(_context.InitiatingUserId);
            _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
        }

        /// <summary>
        /// Handles the GrantAccess message to create a sales team record if it doesn't exist.
        /// </summary>
        private void HandleGrantAccess()
        {
            // Validate input parameters
            if (!_context.InputParameters.Contains("Target") || !_context.InputParameters.Contains("PrincipalAccess"))
            {
                _tracingService.Trace("Missing required input parameters for GrantAccess.");
                return;
            }

            EntityReference target = (EntityReference)_context.InputParameters["Target"];
            PrincipalAccess principalAccess = (PrincipalAccess)_context.InputParameters["PrincipalAccess"];
            EntityReference userOrTeam = principalAccess.Principal;

            // Log details for debugging
            _tracingService.Trace("GrantAccess: User/Team ID=" + userOrTeam.Id + ", LogicalName=" + userOrTeam.LogicalName + ", AccessMask=" + principalAccess.AccessMask + ", Target ID=" + target.Id);

            // Retrieve the target entity
            RetrieveProjectEntity(target);

            // Check for existing sales team record
            EntityCollection existingRecords = GetExistingSalesRecordForCreate(userOrTeam.Id, target);
            _tracingService.Trace("Existing sales team records for GrantAccess: " + existingRecords.Entities.Count);

            if (existingRecords.Entities.Count == 0)
            {
                CreateSalesRecord(userOrTeam.Id, "Grant", target);
            }
        }

        /// <summary>
        /// Handles the RevokeAccess message to update the sales team record's end date.
        /// </summary>
        private void HandleRevokeAccess()
        {
            // Validate input parameters
            if (!_context.InputParameters.Contains("Target") || !_context.InputParameters.Contains("Revokee"))
            {
                _tracingService.Trace("Missing required input parameters for RevokeAccess.");
                return;
            }

            EntityReference target = (EntityReference)_context.InputParameters["Target"];
            EntityReference revokee = (EntityReference)_context.InputParameters["Revokee"];

            // Log details for debugging
            _tracingService.Trace("RevokeAccess: Revokee ID=" + revokee.Id + ", LogicalName=" + revokee.LogicalName + ", Target ID=" + target.Id);

            // Check for existing sales team record
            EntityCollection existingRecords = GetExistingSalesRecord(revokee.Id, target);
            _tracingService.Trace("Existing sales team records for RevokeAccess: " + existingRecords.Entities.Count);

            if (existingRecords.Entities.Count > 0)
            {
                Guid salesTeamId = new Guid(existingRecords.Entities[0]["zox_salesteamid"].ToString());
                UpdateSalesRecord(salesTeamId);
            }
        }

        /// <summary>
        /// Retrieves the project entity based on the target entity reference.
        /// </summary>
        private void RetrieveProjectEntity(EntityReference target)
        {
            if (target == null || string.IsNullOrEmpty(target.LogicalName))
            {
                _tracingService.Trace("Invalid target entity reference.");
                return;
            }

            ColumnSet columns = target.LogicalName == "opportunity" ? _opportunityColumns : _projectColumns;
            _projectEntity = _service.Retrieve(target.LogicalName, target.Id, columns);
        }

        /// <summary>
        /// Retrieves existing sales team records for a user/team and target entity when granting access.
        /// </summary>
        private EntityCollection GetExistingSalesRecordForCreate(Guid userOrTeamId, EntityReference target)
        {
            if (userOrTeamId == Guid.Empty || target == null)
            {
                _tracingService.Trace("Invalid parameters for GetExistingSalesRecordForCreate.");
                return new EntityCollection();
            }

            QueryExpression query = new QueryExpression("zox_salesteam")
            {
                Distinct = true,
                ColumnSet = new ColumnSet("zox_user", "zox_startdate", "zox_enddate", "zox_salesteamid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("zox_user", ConditionOperator.Equal, userOrTeamId),
                        new ConditionExpression("zox_startdate", ConditionOperator.NotNull),
                        new ConditionExpression("zox_enddate", ConditionOperator.Null)
                    }
                },
                TopCount = 1,
                Orders = { new OrderExpression("zox_startdate", OrderType.Descending) }
            };

            AddEntityCondition(query, target);
            return _service.RetrieveMultiple(query);
        }

        /// <summary>
        /// Retrieves existing sales team records for a revokee and target entity when revoking access.
        /// </summary>
        private EntityCollection GetExistingSalesRecord(Guid revokeeId, EntityReference target)
        {
            if (revokeeId == Guid.Empty || target == null)
            {
                _tracingService.Trace("Invalid parameters for GetExistingSalesRecord.");
                return new EntityCollection();
            }

            QueryExpression query = new QueryExpression("zox_salesteam")
            {
                ColumnSet = new ColumnSet("zox_user", "zox_startdate", "zox_salesteamid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("zox_user", ConditionOperator.Equal, revokeeId),
                        new ConditionExpression("zox_startdate", ConditionOperator.NotNull)
                    }
                },
                TopCount = 1,
                Orders = { new OrderExpression("zox_startdate", OrderType.Descending) }
            };

            AddEntityCondition(query, target);
            return _service.RetrieveMultiple(query);
        }

        /// <summary>
        /// Adds entity-specific conditions to the query based on the target entity type.
        /// </summary>
        private void AddEntityCondition(QueryExpression query, EntityReference target)
        {
            if (target.LogicalName == "zox_prelead")
            {
                query.Criteria.AddCondition("zox_prelead", ConditionOperator.Equal, target.Id);
            }
            else if (target.LogicalName == "lead")
            {
                query.Criteria.AddCondition("zox_lead", ConditionOperator.Equal, target.Id);
            }
            else if (target.LogicalName == "opportunity")
            {
                query.Criteria.AddCondition("zox_opportunity", ConditionOperator.Equal, target.Id);
            }
        }

        /// <summary>
        /// Creates a new sales team record for the specified user/team and target entity.
        /// </summary>
        private void CreateSalesRecord(Guid userOrTeamId, string action, EntityReference target)
        {
            if (userOrTeamId == Guid.Empty || target == null || string.IsNullOrEmpty(action))
            {
                _tracingService.Trace("Invalid parameters for CreateSalesRecord.");
                return;
            }

            Entity salesTeam = new Entity("zox_salesteam");
            salesTeam["zox_user"] = new EntityReference("systemuser", userOrTeamId);

            // Set start or end date based on action
            if (action == "Grant")
            {
                salesTeam["zox_startdate"] = DateTime.Now;
            }
            else if (action == "Revoke")
            {
                salesTeam["zox_enddate"] = DateTime.Now;
            }

            // Set entity-specific fields
            SetEntitySpecificFields(salesTeam, target);

            // Retrieve and set user details
            Entity userDetail = _service.Retrieve("systemuser", userOrTeamId, _userColumns);
            if (userDetail.Attributes.Contains("zox_role"))
            {
                salesTeam["zox_role"] = new OptionSetValue(((OptionSetValue)userDetail["zox_role"]).Value);
            }
            if (userDetail.Attributes.Contains("zox_lob"))
            {
                salesTeam["zox_lob"] = new OptionSetValue(((OptionSetValue)userDetail["zox_lob"]).Value);
            }

            _service.Create(salesTeam);
            _tracingService.Trace("Created sales team record for User/Team ID=" + userOrTeamId + ", Target ID=" + target.Id);
        }

        /// <summary>
        /// Sets entity-specific fields (PreLead, Lead, Opportunity, Project, Package) for the sales team record.
        /// </summary>
        private void SetEntitySpecificFields(Entity salesTeam, EntityReference target)
        {
            if (target.LogicalName == "zox_prelead")
            {
                salesTeam["zox_prelead"] = target;
            }
            else if (target.LogicalName == "lead")
            {
                salesTeam["zox_lead"] = target;
            }
            else if (target.LogicalName == "opportunity")
            {
                salesTeam["zox_opportunity"] = target;
                if (_projectEntity != null && _projectEntity.Contains("zox_projecthierarchy") && _projectEntity["zox_projecthierarchy"] != null)
                {
                    salesTeam["zox_project"] = new EntityReference("zox_project", ((EntityReference)_projectEntity["zox_projecthierarchy"]).Id);
                }
            }

            // Set project and package references
            if (_projectEntity != null)
            {
                if (_projectEntity.Contains("zox_project") && _projectEntity["zox_project"] != null)
                {
                    salesTeam["zox_project"] = new EntityReference("zox_project", ((EntityReference)_projectEntity["zox_project"]).Id);
                }
                if (_projectEntity.Contains("zox_package") && _projectEntity["zox_package"] != null)
                {
                    salesTeam["zox_package"] = new EntityReference("zox_project", ((EntityReference)_projectEntity["zox_package"]).Id);
                }
            }
        }

        /// <summary>
        /// Updates the end date of an existing sales team record.
        /// </summary>
        private void UpdateSalesRecord(Guid recordId)
        {
            if (recordId == Guid.Empty)
            {
                _tracingService.Trace("Invalid sales team record ID for update.");
                return;
            }

            Entity salesTeam = new Entity("zox_salesteam");
            salesTeam["zox_salesteamid"] = recordId;
            salesTeam["zox_enddate"] = DateTime.Now;

            _service.Update(salesTeam);
            _tracingService.Trace("Updated end date for sales team record ID=" + recordId);
        }
    }
}