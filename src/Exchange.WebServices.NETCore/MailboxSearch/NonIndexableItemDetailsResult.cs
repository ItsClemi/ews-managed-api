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
///     Represents non indexable item details result.
/// </summary>
[PublicAPI]
public sealed class NonIndexableItemDetailsResult
{
    /// <summary>
    ///     Collection of items
    /// </summary>
    public NonIndexableItem[]? Items { get; set; }

    /// <summary>
    ///     Failed mailboxes
    /// </summary>
    public FailedSearchMailbox[]? FailedMailboxes { get; set; }

    /// <summary>
    ///     Load from xml
    /// </summary>
    /// <param name="reader">The reader</param>
    /// <returns>Non indexable item details result object</returns>
    internal static NonIndexableItemDetailsResult LoadFromXml(EwsServiceXmlReader reader)
    {
        var nonIndexableItemDetailsResult = new NonIndexableItemDetailsResult();
        reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.NonIndexableItemDetailsResult);

        do
        {
            reader.Read();

            if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Items))
            {
                var nonIndexableItems = new List<NonIndexableItem>();
                if (!reader.IsEmptyElement)
                {
                    do
                    {
                        reader.Read();
                        var nonIndexableItem = NonIndexableItem.LoadFromXml(reader);
                        if (nonIndexableItem != null)
                        {
                            nonIndexableItems.Add(nonIndexableItem);
                        }
                    } while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.Items));

                    nonIndexableItemDetailsResult.Items = nonIndexableItems.ToArray();
                }
            }

            if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.FailedMailboxes))
            {
                nonIndexableItemDetailsResult.FailedMailboxes =
                    FailedSearchMailbox.LoadFailedMailboxesXml(XmlNamespace.Types, reader);
            }
        } while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.NonIndexableItemDetailsResult));

        return nonIndexableItemDetailsResult;
    }
}
