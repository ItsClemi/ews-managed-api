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
///     Represents the GetSearchableMailboxes response.
/// </summary>
[PublicAPI]
public sealed class GetSearchableMailboxesResponse : ServiceResponse
{
    readonly List<SearchableMailbox> _searchableMailboxes = new List<SearchableMailbox>();

    /// <summary>
    ///     Searchable mailboxes result
    /// </summary>
    public SearchableMailbox[] SearchableMailboxes => _searchableMailboxes.ToArray();

    /// <summary>
    ///     Failed mailboxes
    /// </summary>
    public FailedSearchMailbox[] FailedMailboxes { get; set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="GetSearchableMailboxesResponse" /> class.
    /// </summary>
    internal GetSearchableMailboxesResponse()
    {
    }

    /// <summary>
    ///     Reads response elements from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
    {
        _searchableMailboxes.Clear();

        base.ReadElementsFromXml(reader);

        reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.SearchableMailboxes);
        if (!reader.IsEmptyElement)
        {
            do
            {
                reader.Read();
                if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.SearchableMailbox))
                {
                    _searchableMailboxes.Add(SearchableMailbox.LoadFromXml(reader));
                }
            } while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.SearchableMailboxes));
        }

        reader.Read();
        if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.FailedMailboxes))
        {
            FailedMailboxes = FailedSearchMailbox.LoadFailedMailboxesXml(XmlNamespace.Messages, reader);
        }
    }
}
