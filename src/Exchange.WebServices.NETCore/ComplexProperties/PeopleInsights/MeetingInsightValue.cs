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
///     Represents the MeetingInsightValue.
/// </summary>
[PublicAPI]
public sealed class MeetingInsightValue : InsightValue
{
    /// <summary>
    ///     Gets the Id
    /// </summary>
    public string Id { get; internal set; }

    /// <summary>
    ///     Gets the Subject
    /// </summary>
    public string Subject { get; internal set; }

    /// <summary>
    ///     Gets the StartUtcTicks
    /// </summary>
    public long StartUtcTicks { get; internal set; }

    /// <summary>
    ///     Gets the EndUtcTicks
    /// </summary>
    public long EndUtcTicks { get; internal set; }

    /// <summary>
    ///     Gets the Location
    /// </summary>
    public string Location { get; internal set; }

    /// <summary>
    ///     Gets the Organizer
    /// </summary>
    public ProfileInsightValue Organizer { get; internal set; }

    /// <summary>
    ///     Gets the Attendees
    /// </summary>
    public ProfileInsightValueCollection Attendees { get; internal set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="MeetingInsightValue" /> class.
    /// </summary>
    public MeetingInsightValue()
    {
        Attendees = new ProfileInsightValueCollection();
    }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">XML reader</param>
    /// <returns>Whether the element was read</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.InsightSource:
            {
                InsightSource = reader.ReadElementValue<string>();
                break;
            }
            case XmlElementNames.UpdatedUtcTicks:
            {
                UpdatedUtcTicks = reader.ReadElementValue<long>();
                break;
            }
            case XmlElementNames.Id:
            {
                Id = reader.ReadElementValue();
                break;
            }
            case XmlElementNames.Subject:
            {
                Subject = reader.ReadElementValue();
                break;
            }
            case XmlElementNames.StartUtcTicks:
            {
                StartUtcTicks = reader.ReadElementValue<long>();
                break;
            }
            case XmlElementNames.EndUtcTicks:
            {
                EndUtcTicks = reader.ReadElementValue<long>();
                break;
            }
            case XmlElementNames.Location:
            {
                Location = reader.ReadElementValue();
                break;
            }
            case XmlElementNames.Organizer:
            {
                Organizer = new ProfileInsightValue();
                Organizer.LoadFromXml(reader, reader.LocalName);
                break;
            }
            case XmlElementNames.Attendees:
            {
                Attendees = new ProfileInsightValueCollection(XmlElementNames.Item);
                Attendees.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.Attendees);
                break;
            }
            default:
            {
                return false;
            }
        }

        return true;
    }
}
