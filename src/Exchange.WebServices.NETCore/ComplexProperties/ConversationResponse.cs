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

[PublicAPI]
public sealed class ConversationResponse : ComplexProperty
{
    /// <summary>
    ///     Property set used to fetch items in the conversation.
    /// </summary>
    private readonly PropertySet _propertySet;

    /// <summary>
    ///     Gets the conversation id.
    /// </summary>
    public ConversationId ConversationId { get; internal set; }

    /// <summary>
    ///     Gets the sync state.
    /// </summary>
    public string SyncState { get; internal set; }

    /// <summary>
    ///     Gets the conversation nodes.
    /// </summary>
    public ConversationNodeCollection ConversationNodes { get; internal set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ConversationResponse" /> class.
    /// </summary>
    /// <param name="propertySet">The property set.</param>
    internal ConversationResponse(PropertySet propertySet)
    {
        _propertySet = propertySet;
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
            case XmlElementNames.ConversationId:
            {
                ConversationId = new ConversationId();
                ConversationId.LoadFromXml(reader, XmlElementNames.ConversationId);
                return true;
            }
            case XmlElementNames.SyncState:
            {
                SyncState = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.ConversationNodes:
            {
                ConversationNodes = new ConversationNodeCollection(_propertySet);
                ConversationNodes.LoadFromXml(reader, XmlElementNames.ConversationNodes);
                return true;
            }
            default:
            {
                return false;
            }
        }
    }
}
