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
///     Represents the schema for meeting messages.
/// </summary>
[PublicAPI]
[Schema]
public class MeetingMessageSchema : EmailMessageSchema
{
    /// <summary>
    ///     Field URIs for MeetingMessage.
    /// </summary>
    private static class FieldUris
    {
        public const string AssociatedCalendarItemId = "meeting:AssociatedCalendarItemId";
        public const string IsDelegated = "meeting:IsDelegated";
        public const string IsOutOfDate = "meeting:IsOutOfDate";
        public const string HasBeenProcessed = "meeting:HasBeenProcessed";
        public const string ResponseType = "meeting:ResponseType";
        public const string IsOrganizer = "cal:IsOrganizer";
    }

    /// <summary>
    ///     Defines the AssociatedAppointmentId property.
    /// </summary>
    public static readonly PropertyDefinition AssociatedAppointmentId = new ComplexPropertyDefinition<ItemId>(
        XmlElementNames.AssociatedCalendarItemId,
        FieldUris.AssociatedCalendarItemId,
        ExchangeVersion.Exchange2007_SP1,
        () => new ItemId()
    );

    /// <summary>
    ///     Defines the IsDelegated property.
    /// </summary>
    public static readonly PropertyDefinition IsDelegated = new BoolPropertyDefinition(
        XmlElementNames.IsDelegated,
        FieldUris.IsDelegated,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the IsOutOfDate property.
    /// </summary>
    public static readonly PropertyDefinition IsOutOfDate = new BoolPropertyDefinition(
        XmlElementNames.IsOutOfDate,
        FieldUris.IsOutOfDate,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the HasBeenProcessed property.
    /// </summary>
    public static readonly PropertyDefinition HasBeenProcessed = new BoolPropertyDefinition(
        XmlElementNames.HasBeenProcessed,
        FieldUris.HasBeenProcessed,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the ResponseType property.
    /// </summary>
    public static readonly PropertyDefinition ResponseType = new GenericPropertyDefinition<MeetingResponseType>(
        XmlElementNames.ResponseType,
        FieldUris.ResponseType,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2007_SP1
    );

    /// <summary>
    ///     Defines the iCalendar Uid property.
    /// </summary>
    public static readonly PropertyDefinition ICalUid = AppointmentSchema.ICalUid;

    /// <summary>
    ///     Defines the iCalendar RecurrenceId property.
    /// </summary>
    public static readonly PropertyDefinition ICalRecurrenceId = AppointmentSchema.ICalRecurrenceId;

    /// <summary>
    ///     Defines the iCalendar DateTimeStamp property.
    /// </summary>
    public static readonly PropertyDefinition ICalDateTimeStamp = AppointmentSchema.ICalDateTimeStamp;

    /// <summary>
    ///     Defines the IsOrganizer property.
    /// </summary>
    public static readonly PropertyDefinition IsOrganizer = new GenericPropertyDefinition<bool>(
        XmlElementNames.IsOrganizer,
        FieldUris.IsOrganizer,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2013
    );

    // This must be after the declaration of property definitions
    internal new static readonly MeetingMessageSchema Instance = new();

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

        RegisterProperty(AssociatedAppointmentId);
        RegisterProperty(IsDelegated);
        RegisterProperty(IsOutOfDate);
        RegisterProperty(HasBeenProcessed);
        RegisterProperty(ResponseType);
        RegisterProperty(ICalUid);
        RegisterProperty(ICalRecurrenceId);
        RegisterProperty(ICalDateTimeStamp);
        RegisterProperty(IsOrganizer);
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="MeetingMessageSchema" /> class.
    /// </summary>
    internal MeetingMessageSchema()
    {
    }
}
