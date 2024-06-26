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
///     Represents a time zone period as defined in the EWS schema.
/// </summary>
internal class TimeZonePeriod : ComplexProperty
{
    internal const string StandardPeriodId = "Std";
    internal const string StandardPeriodName = "Standard";
    internal const string DaylightPeriodId = "Dlt";
    internal const string DaylightPeriodName = "Daylight";

    /// <summary>
    ///     Gets a value indicating whether this period represents the Standard period.
    /// </summary>
    /// <value>
    ///     <c>true</c> if this instance is standard period; otherwise, <c>false</c>.
    /// </value>
    internal bool IsStandardPeriod => string.Compare(Name, StandardPeriodName, StringComparison.OrdinalIgnoreCase) == 0;

    /// <summary>
    ///     Gets or sets the bias to UTC associated with this period.
    /// </summary>
    internal TimeSpan Bias { get; set; }

    /// <summary>
    ///     Gets or sets the name of this period.
    /// </summary>
    internal string Name { get; set; }

    /// <summary>
    ///     Gets or sets the id of this period.
    /// </summary>
    internal string Id { get; set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="TimeZonePeriod" /> class.
    /// </summary>
    internal TimeZonePeriod()
    {
    }

    /// <summary>
    ///     Reads the attributes from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
    {
        Id = reader.ReadAttributeValue(XmlAttributeNames.Id);
        Name = reader.ReadAttributeValue(XmlAttributeNames.Name);
        Bias = EwsUtilities.XsDurationToTimeSpan(reader.ReadAttributeValue(XmlAttributeNames.Bias));
    }

    /// <summary>
    ///     Writes the attributes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteAttributeValue(XmlAttributeNames.Bias, EwsUtilities.TimeSpanToXsDuration(Bias));
        writer.WriteAttributeValue(XmlAttributeNames.Name, Name);
        writer.WriteAttributeValue(XmlAttributeNames.Id, Id);
    }

    /// <summary>
    ///     Loads from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal void LoadFromXml(EwsServiceXmlReader reader)
    {
        LoadFromXml(reader, XmlElementNames.Period);
    }

    /// <summary>
    ///     Writes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal void WriteToXml(EwsServiceXmlWriter writer)
    {
        WriteToXml(writer, XmlElementNames.Period);
    }
}
