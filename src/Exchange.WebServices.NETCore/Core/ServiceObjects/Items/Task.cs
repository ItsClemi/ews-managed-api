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
///     Represents a Task item. Properties available on tasks are defined in the TaskSchema class.
/// </summary>
[PublicAPI]
[Attachable]
[ServiceObjectDefinition(XmlElementNames.Task)]
public class Task : Item
{
    /// <summary>
    ///     Initializes an unsaved local instance of <see cref="Task" />. To bind to an existing task, use Task.Bind() instead.
    /// </summary>
    /// <param name="service">The ExchangeService instance to which this task is bound.</param>
    public Task(ExchangeService service)
        : base(service)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="Task" /> class.
    /// </summary>
    /// <param name="parentAttachment">The parent attachment.</param>
    internal Task(ItemAttachment parentAttachment)
        : base(parentAttachment)
    {
    }

    /// <summary>
    ///     Binds to an existing task and loads the specified set of properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to use to bind to the task.</param>
    /// <param name="id">The Id of the task to bind to.</param>
    /// <param name="propertySet">The set of properties to load.</param>
    /// <param name="token"></param>
    /// <returns>A Task instance representing the task corresponding to the specified Id.</returns>
    public new static Task<Task> Bind(
        ExchangeService service,
        ItemId id,
        PropertySet propertySet,
        CancellationToken token = default
    )
    {
        return service.BindToItem<Task>(id, propertySet, token);
    }

    /// <summary>
    ///     Binds to an existing task and loads its first class properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to use to bind to the task.</param>
    /// <param name="id">The Id of the task to bind to.</param>
    /// <returns>A Task instance representing the task corresponding to the specified Id.</returns>
    public new static Task<Task> Bind(ExchangeService service, ItemId id)
    {
        return Bind(service, id, PropertySet.FirstClassProperties);
    }

    /// <summary>
    ///     Internal method to return the schema associated with this type of object.
    /// </summary>
    /// <returns>The schema associated with this type of object.</returns>
    internal override ServiceObjectSchema GetSchema()
    {
        return TaskSchema.Instance;
    }

    /// <summary>
    ///     Gets the minimum required server version.
    /// </summary>
    /// <returns>Earliest Exchange version in which this service object type is supported.</returns>
    internal override ExchangeVersion GetMinimumRequiredServerVersion()
    {
        return ExchangeVersion.Exchange2007_SP1;
    }

    /// <summary>
    ///     Gets a value indicating whether a time zone SOAP header should be emitted in a CreateItem
    ///     or UpdateItem request so this item can be property saved or updated.
    /// </summary>
    /// <param name="isUpdateOperation">Indicates whether the operation being performed is an update operation.</param>
    /// <returns>
    ///     <c>true</c> if a time zone SOAP header should be emitted; otherwise, <c>false</c>.
    /// </returns>
    internal override bool GetIsTimeZoneHeaderRequired(bool isUpdateOperation)
    {
        return true;
    }

    /// <summary>
    ///     Deletes the current occurrence of a recurring task. After the current occurrence isdeleted,
    ///     the task represents the next occurrence. Developers should call Load to retrieve the new property
    ///     values of the task. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="deleteMode">The deletion mode.</param>
    /// <param name="token"></param>
    public Task<ServiceResponseCollection<ServiceResponse>> DeleteCurrentOccurrence(
        DeleteMode deleteMode,
        CancellationToken token = default
    )
    {
        return InternalDelete(deleteMode, null, AffectedTaskOccurrence.SpecifiedOccurrenceOnly, token);
    }

    /// <summary>
    ///     Applies the local changes that have been made to this task. Calling this method results in at least one call to
    ///     EWS.
    ///     Mutliple calls to EWS might be made if attachments have been added or removed.
    /// </summary>
    /// <param name="conflictResolutionMode">Specifies how conflicts should be resolved.</param>
    /// <param name="token"></param>
    /// <returns>
    ///     A Task object representing the completed occurrence if the task is recurring and the update marks it as completed;
    ///     or
    ///     a Task object representing the current occurrence if the task is recurring and the uypdate changed its recurrence
    ///     pattern; or null in every other case.
    /// </returns>
    public new async Task<Task?> Update(
        ConflictResolutionMode conflictResolutionMode,
        CancellationToken token = default
    )
    {
        return (Task?)await InternalUpdate(
                null /* parentFolder */,
                conflictResolutionMode,
                MessageDisposition.SaveOnly,
                null,
                token
            )
            .ConfigureAwait(false);
    }


    #region Properties

    /// <summary>
    ///     Gets or sets the actual amount of time that is spent on the task.
    /// </summary>
    public int ActualWork
    {
        get => (int)PropertyBag[TaskSchema.ActualWork];
        set => PropertyBag[TaskSchema.ActualWork] = value;
    }

    /// <summary>
    ///     Gets the date and time the task was assigned.
    /// </summary>
    public DateTime AssignedTime => (DateTime)PropertyBag[TaskSchema.AssignedTime];

    /// <summary>
    ///     Gets or sets the billing information of the task.
    /// </summary>
    public string BillingInformation
    {
        get => (string)PropertyBag[TaskSchema.BillingInformation];
        set => PropertyBag[TaskSchema.BillingInformation] = value;
    }

    /// <summary>
    ///     Gets the number of times the task has changed since it was created.
    /// </summary>
    public int ChangeCount => (int)PropertyBag[TaskSchema.ChangeCount];

    /// <summary>
    ///     Gets or sets a list of companies associated with the task.
    /// </summary>
    public StringList Companies
    {
        get => (StringList)PropertyBag[TaskSchema.Companies];
        set => PropertyBag[TaskSchema.Companies] = value;
    }

    /// <summary>
    ///     Gets or sets the date and time on which the task was completed.
    /// </summary>
    public DateTime CompleteDate
    {
        get => (DateTime)PropertyBag[TaskSchema.CompleteDate];
        set => PropertyBag[TaskSchema.CompleteDate] = value;
    }

    /// <summary>
    ///     Gets or sets a list of contacts associated with the task.
    /// </summary>
    public StringList Contacts
    {
        get => (StringList)PropertyBag[TaskSchema.Contacts];
        set => PropertyBag[TaskSchema.Contacts] = value;
    }

    /// <summary>
    ///     Gets the current delegation state of the task.
    /// </summary>
    public TaskDelegationState DelegationState => (TaskDelegationState)PropertyBag[TaskSchema.DelegationState];

    /// <summary>
    ///     Gets the name of the delegator of this task.
    /// </summary>
    public string Delegator => (string)PropertyBag[TaskSchema.Delegator];

    /// <summary>
    ///     Gets or sets the date and time on which the task is due.
    /// </summary>
    public DateTime DueDate
    {
        get => (DateTime)PropertyBag[TaskSchema.DueDate];
        set => PropertyBag[TaskSchema.DueDate] = value;
    }

    /// <summary>
    ///     Gets a value indicating the mode of the task.
    /// </summary>
    public TaskMode Mode => (TaskMode)PropertyBag[TaskSchema.Mode];

    /// <summary>
    ///     Gets a value indicating whether the task is complete.
    /// </summary>
    public bool IsComplete => (bool)PropertyBag[TaskSchema.IsComplete];

    /// <summary>
    ///     Gets a value indicating whether the task is recurring.
    /// </summary>
    public bool IsRecurring => (bool)PropertyBag[TaskSchema.IsRecurring];

    /// <summary>
    ///     Gets a value indicating whether the task is a team task.
    /// </summary>
    public bool IsTeamTask => (bool)PropertyBag[TaskSchema.IsTeamTask];

    /// <summary>
    ///     Gets or sets the mileage of the task.
    /// </summary>
    public string Mileage
    {
        get => (string)PropertyBag[TaskSchema.Mileage];
        set => PropertyBag[TaskSchema.Mileage] = value;
    }

    /// <summary>
    ///     Gets the name of the owner of the task.
    /// </summary>
    public string Owner => (string)PropertyBag[TaskSchema.Owner];

    /// <summary>
    ///     Gets or sets the completeion percentage of the task. PercentComplete must be between 0 and 100.
    /// </summary>
    public double PercentComplete
    {
        get => (double)PropertyBag[TaskSchema.PercentComplete];
        set => PropertyBag[TaskSchema.PercentComplete] = value;
    }

    /// <summary>
    ///     Gets or sets the recurrence pattern for this task. Available recurrence pattern classes include
    ///     Recurrence.DailyPattern, Recurrence.MonthlyPattern and Recurrence.YearlyPattern.
    /// </summary>
    public Recurrence Recurrence
    {
        get => (Recurrence)PropertyBag[TaskSchema.Recurrence];
        set => PropertyBag[TaskSchema.Recurrence] = value;
    }

    /// <summary>
    ///     Gets or sets the date and time on which the task starts.
    /// </summary>
    public DateTime StartDate
    {
        get => (DateTime)PropertyBag[TaskSchema.StartDate];
        set => PropertyBag[TaskSchema.StartDate] = value;
    }

    /// <summary>
    ///     Gets or sets the status of the task.
    /// </summary>
    public TaskStatus Status
    {
        get => (TaskStatus)PropertyBag[TaskSchema.Status];
        set => PropertyBag[TaskSchema.Status] = value;
    }

    /// <summary>
    ///     Gets a string representing the status of the task, localized according to the PreferredCulture
    ///     property of the ExchangeService object the task is bound to.
    /// </summary>
    public string StatusDescription => (string)PropertyBag[TaskSchema.StatusDescription];

    /// <summary>
    ///     Gets or sets the total amount of work spent on the task.
    /// </summary>
    public int TotalWork
    {
        get => (int)PropertyBag[TaskSchema.TotalWork];
        set => PropertyBag[TaskSchema.TotalWork] = value;
    }

    /// <summary>
    ///     Gets the default setting for how to treat affected task occurrences on Delete.
    /// </summary>
    /// <value>AffectedTaskOccurrence.AllOccurrences: All affected Task occurrences will be deleted.</value>
    internal override AffectedTaskOccurrence? DefaultAffectedTaskOccurrences => AffectedTaskOccurrence.AllOccurrences;

    #endregion
}
