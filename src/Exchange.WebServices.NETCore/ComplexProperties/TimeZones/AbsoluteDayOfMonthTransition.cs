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

using Microsoft.Exchange.WebServices.Data.Misc;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a time zone period transition that occurs on a specific day of a specific month.
/// </summary>
internal class AbsoluteDayOfMonthTransition : AbsoluteMonthTransition
{
    /// <summary>
    ///     Gets the day of then month when this transition occurs.
    /// </summary>
    internal int DayOfMonth { get; private set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="AbsoluteDayOfMonthTransition" /> class.
    /// </summary>
    /// <param name="timeZoneDefinition">The time zone definition this transition belongs to.</param>
    internal AbsoluteDayOfMonthTransition(TimeZoneDefinition timeZoneDefinition)
        : base(timeZoneDefinition)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="AbsoluteDayOfMonthTransition" /> class.
    /// </summary>
    /// <param name="timeZoneDefinition">The time zone definition this transition belongs to.</param>
    /// <param name="targetPeriod">The period the transition will target.</param>
    internal AbsoluteDayOfMonthTransition(TimeZoneDefinition timeZoneDefinition, TimeZonePeriod targetPeriod)
        : base(timeZoneDefinition, targetPeriod)
    {
    }

    /// <summary>
    ///     Gets the XML element name associated with the transition.
    /// </summary>
    /// <returns>The XML element name associated with the transition.</returns>
    internal override string GetXmlElementName()
    {
        return XmlElementNames.RecurringDateTransition;
    }

    /// <summary>
    ///     Creates a timw zone transition time.
    /// </summary>
    /// <returns>A TimeZoneInfo.TransitionTime.</returns>
    internal override TransitionTime CreateTransitionTime()
    {
        return TransitionTime.CreateFixedDateRule(new DateTime(TimeOffset.Ticks), Month, DayOfMonth);
    }

    /// <summary>
    ///     Initializes this transition based on the specified transition time.
    /// </summary>
    /// <param name="transitionTime">The transition time to initialize from.</param>
    internal override void InitializeFromTransitionTime(TransitionTime transitionTime)
    {
        base.InitializeFromTransitionTime(transitionTime);

        DayOfMonth = transitionTime.Day;
    }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        if (base.TryReadElementFromXml(reader))
        {
            return true;
        }

        if (reader.LocalName == XmlElementNames.Day)
        {
            DayOfMonth = reader.ReadElementValue<int>();

            EwsUtilities.Assert(
                DayOfMonth > 0 && DayOfMonth <= 31,
                "AbsoluteDayOfMonthTransition.TryReadElementFromXml",
                "dayOfMonth is not in the valid 1 - 31 range."
            );

            return true;
        }

        return false;
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        base.WriteElementsToXml(writer);

        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Day, DayOfMonth);
    }
}
