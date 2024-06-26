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
///     Represents the MIME content of an item.
/// </summary>
[PublicAPI]
public abstract class MimeContentBase : ComplexProperty
{
    /// <summary>
    ///     characterSet returned
    /// </summary>
    private string? _characterSet;

    /// <summary>
    ///     content received
    /// </summary>
    private byte[]? _content;

    /// <summary>
    ///     Gets or sets the character set of the content.
    /// </summary>
    public string? CharacterSet
    {
        get => _characterSet;
        set => SetFieldValue(ref _characterSet, value);
    }

    /// <summary>
    ///     Gets or sets the content.
    /// </summary>
    public byte[]? Content
    {
        get => _content;
        set => SetFieldValue(ref _content, value);
    }

    /// <summary>
    ///     Reads attributes from XML.
    ///     This should always be UTF-8 for MimeContentUTF8
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
    {
        _characterSet = reader.ReadAttributeValue<string>(XmlAttributeNames.CharacterSet);
    }

    /// <summary>
    ///     Reads text value from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadTextValueFromXml(EwsServiceXmlReader reader)
    {
        _content = Convert.FromBase64String(reader.ReadValue());
    }

    /// <summary>
    ///     Writes attributes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteAttributeValue(XmlAttributeNames.CharacterSet, CharacterSet);
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        if (Content != null && Content.Length > 0)
        {
            writer.WriteBase64ElementValue(Content);
        }
    }
}
