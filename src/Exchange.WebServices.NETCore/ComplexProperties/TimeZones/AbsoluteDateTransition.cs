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

using System.Globalization;

using Microsoft.Exchange.WebServices.Data.Misc;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a time zone period transition that occurs on a fixed (absolute) date.
/// </summary>
internal class AbsoluteDateTransition : TimeZoneTransition
{
    /// <summary>
    ///     Gets or sets the absolute date and time when the transition occurs.
    /// </summary>
    internal DateTime DateTime { get; set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="AbsoluteDateTransition" /> class.
    /// </summary>
    /// <param name="timeZoneDefinition">The time zone definition the transition will belong to.</param>
    internal AbsoluteDateTransition(TimeZoneDefinition timeZoneDefinition)
        : base(timeZoneDefinition)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="AbsoluteDateTransition" /> class.
    /// </summary>
    /// <param name="timeZoneDefinition">The time zone definition the transition will belong to.</param>
    /// <param name="targetGroup">The transition group the transition will target.</param>
    internal AbsoluteDateTransition(TimeZoneDefinition timeZoneDefinition, TimeZoneTransitionGroup targetGroup)
        : base(timeZoneDefinition, targetGroup)
    {
    }

    /// <summary>
    ///     Initializes this transition based on the specified transition time.
    /// </summary>
    /// <param name="transitionTime">The transition time to initialize from.</param>
    internal override void InitializeFromTransitionTime(TransitionTime transitionTime)
    {
        throw new ServiceLocalException(Strings.UnsupportedTimeZonePeriodTransitionTarget);
    }

    /// <summary>
    ///     Gets the XML element name associated with the transition.
    /// </summary>
    /// <returns>The XML element name associated with the transition.</returns>
    internal override string GetXmlElementName()
    {
        return XmlElementNames.AbsoluteDateTransition;
    }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        var result = base.TryReadElementFromXml(reader);

        if (!result)
        {
            if (reader.LocalName == XmlElementNames.DateTime)
            {
                DateTime = DateTime.Parse(reader.ReadElementValue(), CultureInfo.InvariantCulture);

                result = true;
            }
        }

        return result;
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        base.WriteElementsToXml(writer);

        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.DateTime, DateTime);
    }
}
