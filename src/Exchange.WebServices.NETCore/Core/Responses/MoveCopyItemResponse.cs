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
///     Represents a response to a Move or Copy operation.
/// </summary>
[PublicAPI]
public sealed class MoveCopyItemResponse : ServiceResponse
{
    /// <summary>
    ///     Gets the copied or moved item. Item is null if the copy or move operation was between
    ///     two mailboxes or between a mailbox and a public folder.
    /// </summary>
    public Item Item { get; private set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="MoveCopyItemResponse" /> class.
    /// </summary>
    internal MoveCopyItemResponse()
    {
    }

    /// <summary>
    ///     Gets Item instance.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <returns>Item.</returns>
    private Item? GetObjectInstance(ExchangeService service, string xmlElementName)
    {
        return EwsUtilities.CreateEwsObjectFromXmlElementName<Item>(service, xmlElementName);
    }

    /// <summary>
    ///     Reads response elements from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
    {
        base.ReadElementsFromXml(reader);

        var items = reader.ReadServiceObjectsCollectionFromXml(
            XmlElementNames.Items,
            GetObjectInstance,
            false,
            null,
            false
        );

        // We only receive the copied or moved items if the copy or move operation was within
        // a single mailbox. No item is returned if the operation is cross-mailbox, from a
        // mailbox to a public folder or from a public folder to a mailbox.
        if (items.Count > 0)
        {
            Item = items[0];
        }
    }
}
