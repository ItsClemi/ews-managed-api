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

using System.Collections.ObjectModel;
using System.Globalization;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a suggestion for a specific date.
/// </summary>
[PublicAPI]
public sealed class Suggestion : ComplexProperty
{
    /// <summary>
    ///     Gets the date and time of the suggestion.
    /// </summary>
    public DateTime Date { get; private set; }

    /// <summary>
    ///     Gets the quality of the suggestion.
    /// </summary>
    public SuggestionQuality Quality { get; private set; }

    /// <summary>
    ///     Gets a collection of suggested times within the suggested day.
    /// </summary>
    public Collection<TimeSuggestion> TimeSuggestions { get; } = new();

    /// <summary>
    ///     Initializes a new instance of the <see cref="Suggestion" /> class.
    /// </summary>
    internal Suggestion()
    {
    }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if appropriate element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.Date:
            {
                // The date that is returned by Availability is unscoped. 
                var tempDate = DateTime.Parse(reader.ReadElementValue(), CultureInfo.InvariantCulture);

                if (tempDate.Kind != DateTimeKind.Unspecified)
                {
                    Date = new DateTime(tempDate.Ticks, DateTimeKind.Unspecified);
                }
                else
                {
                    Date = tempDate;
                }

                return true;
            }
            case XmlElementNames.DayQuality:
            {
                Quality = reader.ReadElementValue<SuggestionQuality>();
                return true;
            }
            case XmlElementNames.SuggestionArray:
            {
                if (!reader.IsEmptyElement)
                {
                    do
                    {
                        reader.Read();

                        if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Suggestion))
                        {
                            var timeSuggestion = new TimeSuggestion();

                            timeSuggestion.LoadFromXml(reader, reader.LocalName);

                            TimeSuggestions.Add(timeSuggestion);
                        }
                    } while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.SuggestionArray));
                }

                return true;
            }
            default:
            {
                return false;
            }
        }
    }
}
