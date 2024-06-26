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
///     Represents recurrence range with start and end dates.
/// </summary>
internal abstract class RecurrenceRange : ComplexProperty
{
    private DateTime _startDate;

    /// <summary>
    ///     Gets the name of the XML element.
    /// </summary>
    /// <value>The name of the XML element.</value>
    internal abstract string XmlElementName { get; }

    /// <summary>
    ///     Gets or sets the recurrence.
    /// </summary>
    /// <value>The recurrence.</value>
    internal Recurrence? Recurrence { get; set; }

    /// <summary>
    ///     Gets or sets the start date.
    /// </summary>
    /// <value>The start date.</value>
    internal DateTime StartDate
    {
        get => _startDate;
        set => SetFieldValue(ref _startDate, value);
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="RecurrenceRange" /> class.
    /// </summary>
    internal RecurrenceRange()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="RecurrenceRange" /> class.
    /// </summary>
    /// <param name="startDate">The start date.</param>
    internal RecurrenceRange(DateTime startDate)
        : this()
    {
        _startDate = startDate;
    }

    /// <summary>
    ///     Changes handler.
    /// </summary>
    internal override void Changed()
    {
        if (Recurrence != null)
        {
            Recurrence.Changed();
        }
    }

    /// <summary>
    ///     Setup the recurrence.
    /// </summary>
    /// <param name="recurrence">The recurrence.</param>
    internal virtual void SetupRecurrence(Recurrence recurrence)
    {
        recurrence.StartDate = StartDate;
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteElementValue(
            XmlNamespace.Types,
            XmlElementNames.StartDate,
            EwsUtilities.DateTimeToXsDate(StartDate)
        );
    }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.StartDate:
            {
                var startDate = reader.ReadElementValueAsUnspecifiedDate();
                if (startDate.HasValue)
                {
                    _startDate = startDate.Value;
                    return true;
                }

                return false;
            }
            default:
            {
                return false;
            }
        }
    }
}
