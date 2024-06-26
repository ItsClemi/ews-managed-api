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
///     Represents the Id of a public folder expressed in a specific format.
/// </summary>
[PublicAPI]
public class AlternatePublicFolderId : AlternateIdBase
{
    /// <summary>
    ///     Name of schema type used for AlternatePublicFolderId element.
    /// </summary>
    internal const string SchemaTypeName = "AlternatePublicFolderIdType";

    /// <summary>
    ///     The Id of the public folder.
    /// </summary>
    public string FolderId { get; set; }

    /// <summary>
    ///     Initializes a new instance of AlternatePublicFolderId.
    /// </summary>
    public AlternatePublicFolderId()
    {
    }

    /// <summary>
    ///     Initializes a new instance of AlternatePublicFolderId.
    /// </summary>
    /// <param name="format">The format in which the public folder Id is expressed.</param>
    /// <param name="folderId">The Id of the public folder.</param>
    public AlternatePublicFolderId(IdFormat format, string folderId)
        : base(format)
    {
        FolderId = folderId;
    }

    /// <summary>
    ///     Gets the name of the XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal override string GetXmlElementName()
    {
        return XmlElementNames.AlternatePublicFolderId;
    }

    /// <summary>
    ///     Writes the attributes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        base.WriteAttributesToXml(writer);

        writer.WriteAttributeValue(XmlAttributeNames.FolderId, FolderId);
    }

    /// <summary>
    ///     Loads the attributes from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void LoadAttributesFromXml(EwsServiceXmlReader reader)
    {
        base.LoadAttributesFromXml(reader);

        FolderId = reader.ReadAttributeValue(XmlAttributeNames.FolderId);
    }
}
