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
///     Represents the response to a GetConversationItems operation.
/// </summary>
[PublicAPI]
public sealed class ConversationNode : ComplexProperty
{
    private readonly PropertySet _propertySet;

    /// <summary>
    ///     Gets or sets the Internet message id of the node.
    /// </summary>
    public string InternetMessageId { get; set; }

    /// <summary>
    ///     Gets or sets the Internet message id of the parent node.
    /// </summary>
    public string ParentInternetMessageId { get; set; }

    /// <summary>
    ///     Gets or sets the items.
    /// </summary>
    public List<Item> Items { get; set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ConversationNode" /> class.
    /// </summary>
    /// <param name="propertySet">The property set.</param>
    internal ConversationNode(PropertySet propertySet)
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
            case XmlElementNames.InternetMessageId:
            {
                InternetMessageId = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.ParentInternetMessageId:
            {
                ParentInternetMessageId = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.Items:
            {
                Items = reader.ReadServiceObjectsCollectionFromXml(
                    XmlNamespace.Types,
                    XmlElementNames.Items,
                    GetObjectInstance,
                    true,
                    _propertySet,
                    false
                );
                return true;
            }

            default:
            {
                return false;
            }
        }
    }

    /// <summary>
    ///     Gets the item instance.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <returns>Item.</returns>
    private Item? GetObjectInstance(ExchangeService service, string xmlElementName)
    {
        return EwsUtilities.CreateEwsObjectFromXmlElementName<Item>(service, xmlElementName);
    }

    /// <summary>
    ///     Gets the name of the XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal static string GetXmlElementName()
    {
        return XmlElementNames.ConversationNode;
    }
}
