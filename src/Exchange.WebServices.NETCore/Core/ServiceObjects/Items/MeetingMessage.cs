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

using System.ComponentModel;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a meeting-related message. Properties available on meeting messages are defined in the
///     MeetingMessageSchema class.
/// </summary>
[PublicAPI]
[ServiceObjectDefinition(XmlElementNames.MeetingMessage)]
[EditorBrowsable(EditorBrowsableState.Never)]
public class MeetingMessage : EmailMessage
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="MeetingMessage" /> class.
    /// </summary>
    /// <param name="parentAttachment">The parent attachment.</param>
    internal MeetingMessage(ItemAttachment parentAttachment)
        : base(parentAttachment)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="MeetingMessage" /> class.
    /// </summary>
    /// <param name="service">EWS service to which this object belongs.</param>
    internal MeetingMessage(ExchangeService service)
        : base(service)
    {
    }

    /// <summary>
    ///     Binds to an existing meeting message and loads the specified set of properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to use to bind to the meeting message.</param>
    /// <param name="id">The Id of the meeting message to bind to.</param>
    /// <param name="propertySet">The set of properties to load.</param>
    /// <param name="token"></param>
    /// <returns>A MeetingMessage instance representing the meeting message corresponding to the specified Id.</returns>
    public new static Task<MeetingMessage> Bind(
        ExchangeService service,
        ItemId id,
        PropertySet propertySet,
        CancellationToken token = default
    )
    {
        return service.BindToItem<MeetingMessage>(id, propertySet, token);
    }

    /// <summary>
    ///     Binds to an existing meeting message and loads its first class properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to use to bind to the meeting message.</param>
    /// <param name="id">The Id of the meeting message to bind to.</param>
    /// <returns>A MeetingMessage instance representing the meeting message corresponding to the specified Id.</returns>
    public new static Task<MeetingMessage> Bind(ExchangeService service, ItemId id)
    {
        return Bind(service, id, PropertySet.FirstClassProperties);
    }

    /// <summary>
    ///     Internal method to return the schema associated with this type of object.
    /// </summary>
    /// <returns>The schema associated with this type of object.</returns>
    internal override ServiceObjectSchema GetSchema()
    {
        return MeetingMessageSchema.Instance;
    }

    /// <summary>
    ///     Gets the minimum required server version.
    /// </summary>
    /// <returns>Earliest Exchange version in which this service object type is supported.</returns>
    internal override ExchangeVersion GetMinimumRequiredServerVersion()
    {
        return ExchangeVersion.Exchange2007_SP1;
    }


    #region Properties

    /// <summary>
    ///     Gets the Id of the appointment associated with the meeting message.
    /// </summary>
    public ItemId AssociatedAppointmentId => (ItemId)PropertyBag[MeetingMessageSchema.AssociatedAppointmentId];

    /// <summary>
    ///     Gets a value indicating whether the meeting message is delegated.
    /// </summary>
    public bool IsDelegated => (bool)PropertyBag[MeetingMessageSchema.IsDelegated];

    /// <summary>
    ///     Gets a value indicating whether the meeting message is out of date.
    /// </summary>
    public bool IsOutOfDate => (bool)PropertyBag[MeetingMessageSchema.IsOutOfDate];

    /// <summary>
    ///     Gets a value indicating whether the meeting message has been processed by Exchange (i.e. Exchange has noted
    ///     the arrival of a meeting request and has created the associated meeting item in the calendar).
    /// </summary>
    public bool HasBeenProcessed => (bool)PropertyBag[MeetingMessageSchema.HasBeenProcessed];

    /// <summary>
    ///     Gets the isorganizer property for this meeting
    /// </summary>
    public bool IsOrganizer => (bool)PropertyBag[MeetingMessageSchema.IsOrganizer];

    /// <summary>
    ///     Gets the type of response the meeting message represents.
    /// </summary>
    public MeetingResponseType ResponseType => (MeetingResponseType)PropertyBag[MeetingMessageSchema.ResponseType];

    /// <summary>
    ///     Gets the ICalendar Uid.
    /// </summary>
    public string ICalUid => (string)PropertyBag[MeetingMessageSchema.ICalUid];

    /// <summary>
    ///     Gets the ICalendar RecurrenceId.
    /// </summary>
    public DateTime ICalRecurrenceId => (DateTime)PropertyBag[MeetingMessageSchema.ICalRecurrenceId];

    /// <summary>
    ///     Gets the ICalendar DateTimeStamp.
    /// </summary>
    public DateTime ICalDateTimeStamp => (DateTime)PropertyBag[MeetingMessageSchema.ICalDateTimeStamp];

    #endregion
}
