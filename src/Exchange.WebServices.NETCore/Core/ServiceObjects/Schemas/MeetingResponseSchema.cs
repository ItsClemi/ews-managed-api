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
public class MeetingResponseSchema : MeetingMessageSchema
{
    /// <summary>
    ///     Field URIs for MeetingMessage.
    /// </summary>
    private static class FieldUris
    {
        public const string ProposedStart = "meeting:ProposedStart";
        public const string ProposedEnd = "meeting:ProposedEnd";
    }

    /// <summary>
    ///     Defines the Start property.
    /// </summary>
    public static readonly PropertyDefinition Start = AppointmentSchema.Start;

    /// <summary>
    ///     Defines the End property.
    /// </summary>
    public static readonly PropertyDefinition End = AppointmentSchema.End;

    /// <summary>
    ///     Defines the Location property.
    /// </summary>
    public static readonly PropertyDefinition Location = AppointmentSchema.Location;

    /// <summary>
    ///     Defines the AppointmentType property.
    /// </summary>
    public static readonly PropertyDefinition AppointmentType = AppointmentSchema.AppointmentType;

    /// <summary>
    ///     Defines the Recurrence property.
    /// </summary>
    public static readonly PropertyDefinition Recurrence = AppointmentSchema.Recurrence;

    /// <summary>
    ///     Defines the Proposed Start property.
    /// </summary>
    public static readonly PropertyDefinition ProposedStart = new ScopedDateTimePropertyDefinition(
        XmlElementNames.ProposedStart,
        FieldUris.ProposedStart,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2013,
        _ => AppointmentSchema.StartTimeZone
    );

    /// <summary>
    ///     Defines the Proposed End property.
    /// </summary>
    public static readonly PropertyDefinition ProposedEnd = new ScopedDateTimePropertyDefinition(
        XmlElementNames.ProposedEnd,
        FieldUris.ProposedEnd,
        PropertyDefinitionFlags.CanFind,
        ExchangeVersion.Exchange2013,
        _ => AppointmentSchema.EndTimeZone
    );

    /// <summary>
    ///     Enhanced Location property.
    /// </summary>
    public static readonly PropertyDefinition EnhancedLocation = AppointmentSchema.EnhancedLocation;

    // This must be after the declaration of property definitions
    internal new static readonly MeetingResponseSchema Instance = new();

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

        RegisterProperty(Start);
        RegisterProperty(End);
        RegisterProperty(Location);
        RegisterProperty(Recurrence);
        RegisterProperty(AppointmentType);
        RegisterProperty(ProposedStart);
        RegisterProperty(ProposedEnd);
        RegisterProperty(EnhancedLocation);
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="MeetingMessageSchema" /> class.
    /// </summary>
    internal MeetingResponseSchema()
    {
    }
}
