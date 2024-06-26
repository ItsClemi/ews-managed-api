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
///     Encapsulates information on the changehighlights of a meeting request.
/// </summary>
[PublicAPI]
public sealed class ChangeHighlights : ComplexProperty
{
    /// <summary>
    ///     Gets a value indicating whether the location has changed.
    /// </summary>
    public bool HasLocationChanged { get; private set; }

    /// <summary>
    ///     Gets the old location
    /// </summary>
    public string Location { get; private set; }

    /// <summary>
    ///     Gets a value indicating whether the the start time has changed.
    /// </summary>
    public bool HasStartTimeChanged { get; private set; }

    /// <summary>
    ///     Gets the old start date and time of the meeting.
    /// </summary>
    public DateTime Start { get; private set; }

    /// <summary>
    ///     Gets a value indicating whether the the end time has changed.
    /// </summary>
    public bool HasEndTimeChanged { get; private set; }

    /// <summary>
    ///     Gets the old end date and time of the meeting.
    /// </summary>
    public DateTime End { get; private set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ChangeHighlights" /> class.
    /// </summary>
    internal ChangeHighlights()
    {
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
            case XmlElementNames.HasLocationChanged:
            {
                HasLocationChanged = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.Location:
            {
                Location = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.HasStartTimeChanged:
            {
                HasStartTimeChanged = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.Start:
            {
                Start = reader.ReadElementValueAsDateTime().Value;
                return true;
            }
            case XmlElementNames.HasEndTimeChanged:
            {
                HasEndTimeChanged = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.End:
            {
                End = reader.ReadElementValueAsDateTime().Value;
                return true;
            }
            default:
            {
                return false;
            }
        }
    }
}
