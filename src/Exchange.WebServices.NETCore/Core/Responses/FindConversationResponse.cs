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
using System.Xml;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents the response to a Conversation search operation.
/// </summary>
internal sealed class FindConversationResponse : ServiceResponse
{
    /// <summary>
    ///     Gets the collection of conversations in results.
    /// </summary>
    internal Collection<Conversation> Conversations => Results.Conversations;

    /// <summary>
    ///     Gets FindConversation results.
    /// </summary>
    /// <returns>FindConversation results.</returns>
    internal FindConversationResults Results { get; private set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="FindConversationResponse" /> class.
    /// </summary>
    internal FindConversationResponse()
    {
        Results = new FindConversationResults();
    }

    /// <summary>
    ///     Read Conversations from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
    {
        EwsUtilities.Assert(
            Results.Conversations != null,
            "FindConversationResponse.ReadElementsFromXml",
            "conversations is null."
        );

        EwsUtilities.Assert(
            Results.HighlightTerms != null,
            "FindConversationResponse.ReadElementsFromXml",
            "highlightTerms is null."
        );

        reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.Conversations);
        if (!reader.IsEmptyElement)
        {
            do
            {
                reader.Read();

                if (reader.NodeType == XmlNodeType.Element)
                {
                    var item = EwsUtilities.CreateEwsObjectFromXmlElementName<Conversation>(
                        reader.Service,
                        reader.LocalName
                    );

                    if (item == null)
                    {
                        reader.SkipCurrentElement();
                    }
                    else
                    {
                        item.LoadFromXml(reader, true, null, false);

                        Results.Conversations.Add(item);
                    }
                }
            } while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.Conversations));
        }

        reader.Read();

        if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.HighlightTerms) && !reader.IsEmptyElement)
        {
            do
            {
                reader.Read();

                if (reader.NodeType == XmlNodeType.Element)
                {
                    var term = new HighlightTerm();

                    term.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.HighlightTerm);

                    Results.HighlightTerms.Add(term);
                }
            } while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.HighlightTerms));
        }

        if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.TotalConversationsInView) &&
            !reader.IsEmptyElement)
        {
            Results.TotalCount = reader.ReadElementValue<int>();

            reader.Read();
        }

        if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.IndexedOffset) && !reader.IsEmptyElement)
        {
            Results.IndexedOffset = reader.ReadElementValue<int>();

            reader.Read();
        }
    }
}
