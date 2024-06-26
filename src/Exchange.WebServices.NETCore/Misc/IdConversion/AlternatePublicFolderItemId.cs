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
///     Represents the Id of a public folder item expressed in a specific format.
/// </summary>
[PublicAPI]
public class AlternatePublicFolderItemId : AlternatePublicFolderId
{
    /// <summary>
    ///     Schema type associated with AlternatePublicFolderItemId.
    /// </summary>
    internal new const string SchemaTypeName = "AlternatePublicFolderItemIdType";

    /// <summary>
    ///     The Id of the public folder item.
    /// </summary>
    public string ItemId { get; set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="AlternatePublicFolderItemId" /> class.
    /// </summary>
    public AlternatePublicFolderItemId()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="AlternatePublicFolderItemId" /> class.
    /// </summary>
    /// <param name="format">The format in which the public folder item Id is expressed.</param>
    /// <param name="folderId">The Id of the parent public folder of the public folder item.</param>
    /// <param name="itemId">The Id of the public folder item.</param>
    public AlternatePublicFolderItemId(IdFormat format, string folderId, string itemId)
        : base(format, folderId)
    {
        ItemId = itemId;
    }

    /// <summary>
    ///     Gets the name of the XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal override string GetXmlElementName()
    {
        return XmlElementNames.AlternatePublicFolderItemId;
    }

    /// <summary>
    ///     Writes the attributes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        base.WriteAttributesToXml(writer);

        writer.WriteAttributeValue(XmlAttributeNames.ItemId, ItemId);
    }

    /// <summary>
    ///     Loads the attributes from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void LoadAttributesFromXml(EwsServiceXmlReader reader)
    {
        base.LoadAttributesFromXml(reader);

        ItemId = reader.ReadAttributeValue(XmlAttributeNames.ItemId);
    }
}
