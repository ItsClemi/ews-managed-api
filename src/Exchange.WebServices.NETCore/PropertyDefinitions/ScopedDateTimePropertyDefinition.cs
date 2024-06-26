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

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Defines a callback method used to get a reference to a property definition.
/// </summary>
/// <param name="version">The EWS version for which the property is to be retrieved.</param>
internal delegate PropertyDefinition GetPropertyDefinitionCallback(ExchangeVersion version);

/// <summary>
///     Represents a property definition for DateTime values scoped to a specific time zone property.
/// </summary>
internal class ScopedDateTimePropertyDefinition : DateTimePropertyDefinition
{
    private readonly GetPropertyDefinitionCallback _getPropertyDefinitionCallback;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ScopedDateTimePropertyDefinition" /> class.
    /// </summary>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <param name="uri">The URI.</param>
    /// <param name="flags">The flags.</param>
    /// <param name="version">The version.</param>
    /// <param name="getPropertyDefinitionCallback">The callback that will be used to retrieve the time zone property.</param>
    internal ScopedDateTimePropertyDefinition(
        string xmlElementName,
        string uri,
        PropertyDefinitionFlags flags,
        ExchangeVersion version,
        GetPropertyDefinitionCallback getPropertyDefinitionCallback
    )
        : base(xmlElementName, uri, flags, version)
    {
        EwsUtilities.Assert(
            getPropertyDefinitionCallback != null,
            "ScopedDateTimePropertyDefinition.ctor",
            "getPropertyDefinitionCallback is null."
        );

        _getPropertyDefinitionCallback = getPropertyDefinitionCallback;
    }

    /// <summary>
    ///     Gets the time zone property to which to scope times.
    /// </summary>
    /// <param name="version">The EWS version for which the property is to be retrieved.</param>
    /// <returns>The PropertyDefinition of the scoping time zone property.</returns>
    private PropertyDefinition GetTimeZoneProperty(ExchangeVersion version)
    {
        var timeZoneProperty = _getPropertyDefinitionCallback(version);

        EwsUtilities.Assert(
            timeZoneProperty != null,
            "ScopedDateTimePropertyDefinition.GetTimeZoneProperty",
            "timeZoneProperty is null."
        );

        return timeZoneProperty;
    }

    /// <summary>
    ///     Scopes the date time property to the appropriate time zone, if necessary.
    /// </summary>
    /// <param name="service">The service emitting the request.</param>
    /// <param name="dateTime">The date time.</param>
    /// <param name="propertyBag">The property bag.</param>
    /// <param name="isUpdateOperation">Indicates whether the scoping is to be performed in the context of an update operation.</param>
    /// <returns>The converted DateTime.</returns>
    internal override DateTime ScopeToTimeZone(
        ExchangeServiceBase service,
        DateTime dateTime,
        PropertyBag propertyBag,
        bool isUpdateOperation
    )
    {
        if (!propertyBag.Owner.GetIsCustomDateTimeScopingRequired())
        {
            // Most item types do not require a custom scoping mechanism. For those item types,
            // use the default scoping mechanism.
            return base.ScopeToTimeZone(service, dateTime, propertyBag, isUpdateOperation);
        }

        // Appointment, however, requires a custom scoping mechanism which is based on an
        // associated time zone property.
        var timeZoneProperty = GetTimeZoneProperty(service.RequestedServerVersion);

        propertyBag.TryGetProperty(timeZoneProperty, out var timeZonePropertyValue);

        if (timeZonePropertyValue != null && propertyBag.IsPropertyUpdated(timeZoneProperty))
        {
            // If we have the associated time zone property handy and if it has been updated locally,
            // then we scope the date time to that time zone.
            try
            {
                var convertedDateTime = EwsUtilities.ConvertTime(
                    dateTime,
                    (TimeZoneInfo)timeZonePropertyValue,
                    TimeZoneInfo.Utc
                );

                // This is necessary to stamp the date/time with the Local kind.
                return new DateTime(convertedDateTime.Ticks, DateTimeKind.Utc);
            }
            catch (TimeZoneConversionException e)
            {
                throw new PropertyException(string.Format(Strings.InvalidDateTime, dateTime), Name, e);
            }
        }

        if (isUpdateOperation)
        {
            // In an update operation, what we do depends on what version of EWS
            // we are targeting.
            if (service.RequestedServerVersion == ExchangeVersion.Exchange2007_SP1)
            {
                // For Exchange 2007 SP1, we still need to scope to the service's time zone.
                return base.ScopeToTimeZone(service, dateTime, propertyBag, isUpdateOperation);
            }

            // Otherwise, we let the server scope to the appropriate time zone.
            return dateTime;
        }

        // In a Create operation, always scope to the service's time zone.
        return base.ScopeToTimeZone(service, dateTime, propertyBag, isUpdateOperation);
    }
}
