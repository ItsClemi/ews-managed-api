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
///     Represents recurrent range with an end date.
/// </summary>
internal sealed class EndDateRecurrenceRange : RecurrenceRange
{
    private DateTime _endDate;

    /// <summary>
    ///     Gets the name of the XML element.
    /// </summary>
    /// <value>The name of the XML element.</value>
    internal override string XmlElementName => XmlElementNames.EndDateRecurrence;

    /// <summary>
    ///     Gets or sets the end date.
    /// </summary>
    /// <value>The end date.</value>
    public DateTime EndDate
    {
        get => _endDate;
        set => SetFieldValue(ref _endDate, value);
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="EndDateRecurrenceRange" /> class.
    /// </summary>
    public EndDateRecurrenceRange()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="EndDateRecurrenceRange" /> class.
    /// </summary>
    /// <param name="startDate">The start date.</param>
    /// <param name="endDate">The end date.</param>
    public EndDateRecurrenceRange(DateTime startDate, DateTime endDate)
        : base(startDate)
    {
        _endDate = endDate;
    }

    /// <summary>
    ///     Setups the recurrence.
    /// </summary>
    /// <param name="recurrence">The recurrence.</param>
    internal override void SetupRecurrence(Recurrence recurrence)
    {
        base.SetupRecurrence(recurrence);

        recurrence.EndDate = EndDate;
    }

    /// <summary>
    ///     Writes the elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        base.WriteElementsToXml(writer);

        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.EndDate, EwsUtilities.DateTimeToXsDate(EndDate));
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

        switch (reader.LocalName)
        {
            case XmlElementNames.EndDate:
            {
                _endDate = reader.ReadElementValueAsDateTime().Value;
                return true;
            }
            default:
            {
                return false;
            }
        }
    }
}
