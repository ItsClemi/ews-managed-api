/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents the schema for task items.
/// </summary>
[PublicAPI]
[Schema]
public class TaskSchema : ItemSchema
{
    /// <summary>
    ///     Field URIs for tasks.
    /// </summary>
    private static class FieldUris
    {
        public const string ActualWork = "task:ActualWork";
        public const string AssignedTime = "task:AssignedTime";
        public const string BillingInformation = "task:BillingInformation";
        public const string ChangeCount = "task:ChangeCount";
        public const string Companies = "task:Companies";
        public const string CompleteDate = "task:CompleteDate";
        public const string Contacts = "task:Contacts";
        public const string DelegationState = "task:DelegationState";
        public const string Delegator = "task:Delegator";
        public const string DueDate = "task:DueDate";
        public const string IsAssignmentEditable = "task:IsAssignmentEditable";
        public const string IsComplete = "task:IsComplete";
        public const string IsRecurring = "task:IsRecurring";
        public const string IsTeamTask = "task:IsTeamTask";
        public const string Mileage = "task:Mileage";
        public const string Owner = "task:Owner";
        public const string PercentComplete = "task:PercentComplete";
        public const string Recurrence = "task:Recurrence";
        public const string StartDate = "task:StartDate";
        public const string Status = "task:Status";
        public const string StatusDescription = "task:StatusDescription";
        public const string TotalWork = "task:TotalWork";
    }

    /// <summary>
    ///     Defines the ActualWork property.
    /// </summary>
    public static readonly PropertyDefinition ActualWork = new IntPropertyDefinition(
        XmlElementNames.ActualWork,
        FieldUris.ActualWork,
        PropertyDefinitionFlags.CanSet |
        PropertyDefinitionFlags.CanUpdate |
        PropertyDefinitionFlags.CanDelete |
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1,
        true
    ); // isNullable

    /// <summary>
    ///     Defines the AssignedTime property.
    /// </summary>
    public static readonly PropertyDefinition AssignedTime = new DateTimePropertyDefinition(
        XmlElementNames.AssignedTime,
        FieldUris.AssignedTime,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1,
        true
    ); // isNullable

    /// <summary>
    ///     Defines the BillingInformation property.
    /// </summary>
    public static readonly PropertyDefinition BillingInformation = new StringPropertyDefinition(
        XmlElementNames.BillingInformation,
        FieldUris.BillingInformation,
        PropertyDefinitionFlags.CanSet |
        PropertyDefinitionFlags.CanUpdate |
        PropertyDefinitionFlags.CanDelete |
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the ChangeCount property.
    /// </summary>
    public static readonly PropertyDefinition ChangeCount = new IntPropertyDefinition(
        XmlElementNames.ChangeCount,
        FieldUris.ChangeCount,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the Companies property.
    /// </summary>
    public static readonly PropertyDefinition Companies = new ComplexPropertyDefinition<StringList>(
        XmlElementNames.Companies,
        FieldUris.Companies,
        PropertyDefinitionFlags.AutoInstantiateOnRead |
        PropertyDefinitionFlags.CanSet |
        PropertyDefinitionFlags.CanUpdate |
        PropertyDefinitionFlags.CanDelete |
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1,
        () => new StringList()
    );

    /// <summary>
    ///     Defines the CompleteDate property.
    /// </summary>
    public static readonly PropertyDefinition CompleteDate = new DateTimePropertyDefinition(
        XmlElementNames.CompleteDate,
        FieldUris.CompleteDate,
        PropertyDefinitionFlags.CanSet |
        PropertyDefinitionFlags.CanUpdate |
        PropertyDefinitionFlags.CanDelete |
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1,
        true
    ); // isNullable

    /// <summary>
    ///     Defines the Contacts property.
    /// </summary>
    public static readonly PropertyDefinition Contacts = new ComplexPropertyDefinition<StringList>(
        XmlElementNames.Contacts,
        FieldUris.Contacts,
        PropertyDefinitionFlags.AutoInstantiateOnRead |
        PropertyDefinitionFlags.CanSet |
        PropertyDefinitionFlags.CanUpdate |
        PropertyDefinitionFlags.CanDelete |
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1,
        () => new StringList()
    );

    /// <summary>
    ///     Defines the DelegationState property.
    /// </summary>
    public static readonly PropertyDefinition DelegationState = new TaskDelegationStatePropertyDefinition(
        XmlElementNames.DelegationState,
        FieldUris.DelegationState,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the Delegator property.
    /// </summary>
    public static readonly PropertyDefinition Delegator = new StringPropertyDefinition(
        XmlElementNames.Delegator,
        FieldUris.Delegator,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the DueDate property.
    /// </summary>
    public static readonly PropertyDefinition DueDate = new DateTimePropertyDefinition(
        XmlElementNames.DueDate,
        FieldUris.DueDate,
        PropertyDefinitionFlags.CanSet |
        PropertyDefinitionFlags.CanUpdate |
        PropertyDefinitionFlags.CanDelete |
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1,
        true
    ); // isNullable

    // TODO : This is the worst possible name for that property

    /// <summary>
    ///     Defines the Mode property.
    /// </summary>
    public static readonly PropertyDefinition Mode = new GenericPropertyDefinition<TaskMode>(
        XmlElementNames.IsAssignmentEditable,
        FieldUris.IsAssignmentEditable,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the IsComplete property.
    /// </summary>
    public static readonly PropertyDefinition IsComplete = new BoolPropertyDefinition(
        XmlElementNames.IsComplete,
        FieldUris.IsComplete,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the IsRecurring property.
    /// </summary>
    public static readonly PropertyDefinition IsRecurring = new BoolPropertyDefinition(
        XmlElementNames.IsRecurring,
        FieldUris.IsRecurring,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the IsTeamTask property.
    /// </summary>
    public static readonly PropertyDefinition IsTeamTask = new BoolPropertyDefinition(
        XmlElementNames.IsTeamTask,
        FieldUris.IsTeamTask,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the Mileage property.
    /// </summary>
    public static readonly PropertyDefinition Mileage = new StringPropertyDefinition(
        XmlElementNames.Mileage,
        FieldUris.Mileage,
        PropertyDefinitionFlags.CanSet |
        PropertyDefinitionFlags.CanUpdate |
        PropertyDefinitionFlags.CanDelete |
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the Owner property.
    /// </summary>
    public static readonly PropertyDefinition Owner = new StringPropertyDefinition(
        XmlElementNames.Owner,
        FieldUris.Owner,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the PercentComplete property.
    /// </summary>
    public static readonly PropertyDefinition PercentComplete = new DoublePropertyDefinition(
        XmlElementNames.PercentComplete,
        FieldUris.PercentComplete,
        PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the Recurrence property.
    /// </summary>
    public static readonly PropertyDefinition Recurrence = new RecurrencePropertyDefinition(
        XmlElementNames.Recurrence,
        FieldUris.Recurrence,
        PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the StartDate property.
    /// </summary>
    public static readonly PropertyDefinition StartDate = new DateTimePropertyDefinition(
        XmlElementNames.StartDate,
        FieldUris.StartDate,
        PropertyDefinitionFlags.CanSet |
        PropertyDefinitionFlags.CanUpdate |
        PropertyDefinitionFlags.CanDelete |
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1,
        true
    ); // isNullable

    /// <summary>
    ///     Defines the Status property.
    /// </summary>
    public static readonly PropertyDefinition Status = new GenericPropertyDefinition<TaskStatus>(
        XmlElementNames.Status,
        FieldUris.Status,
        PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the StatusDescription property.
    /// </summary>
    public static readonly PropertyDefinition StatusDescription = new StringPropertyDefinition(
        XmlElementNames.StatusDescription,
        FieldUris.StatusDescription,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the TotalWork property.
    /// </summary>
    public static readonly PropertyDefinition TotalWork = new IntPropertyDefinition(
        XmlElementNames.TotalWork,
        FieldUris.TotalWork,
        PropertyDefinitionFlags.CanSet |
        PropertyDefinitionFlags.CanUpdate |
        PropertyDefinitionFlags.CanDelete |
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1,
        true
    ); // isNullable

    // This must be declared after the property definitions
    internal new static readonly TaskSchema Instance = new();

    /// <summary>
    ///     Registers properties.
    /// </summary>
    /// <remarks>
    ///     IMPORTANT NOTE: PROPERTIES MUST BE REGISTERED IN SCHEMA ORDER (i.e. the same order as they are defined in
    ///     types.xsd)
    /// </remarks>
    internal override void RegisterProperties()
    {
        base.RegisterProperties();

        RegisterProperty(ActualWork);
        RegisterProperty(AssignedTime);
        RegisterProperty(BillingInformation);
        RegisterProperty(ChangeCount);
        RegisterProperty(Companies);
        RegisterProperty(CompleteDate);
        RegisterProperty(Contacts);
        RegisterProperty(DelegationState);
        RegisterProperty(Delegator);
        RegisterProperty(DueDate);
        RegisterProperty(Mode);
        RegisterProperty(IsComplete);
        RegisterProperty(IsRecurring);
        RegisterProperty(IsTeamTask);
        RegisterProperty(Mileage);
        RegisterProperty(Owner);
        RegisterProperty(PercentComplete);
        RegisterProperty(Recurrence);
        RegisterProperty(StartDate);
        RegisterProperty(Status);
        RegisterProperty(StatusDescription);
        RegisterProperty(TotalWork);
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="TaskSchema" /> class.
    /// </summary>
    internal TaskSchema()
    {
    }
}
