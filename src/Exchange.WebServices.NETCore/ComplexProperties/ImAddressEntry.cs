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

using System.ComponentModel;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents an entry of an ImAddressDictionary.
/// </summary>
[PublicAPI]
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ImAddressEntry : DictionaryEntryProperty<ImAddressKey>
{
    private string _imAddress;

    /// <summary>
    ///     Gets or sets the Instant Messaging address of the entry.
    /// </summary>
    public string ImAddress
    {
        get => _imAddress;
        set => SetFieldValue(ref _imAddress, value);
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ImAddressEntry" /> class.
    /// </summary>
    internal ImAddressEntry()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ImAddressEntry" /> class.
    /// </summary>
    /// <param name="key">The key.</param>
    /// <param name="imAddress">The im address.</param>
    internal ImAddressEntry(ImAddressKey key, string imAddress)
        : base(key)
    {
        _imAddress = imAddress;
    }

    /// <summary>
    ///     Reads the text value from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadTextValueFromXml(EwsServiceXmlReader reader)
    {
        _imAddress = reader.ReadValue();
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteValue(ImAddress, XmlElementNames.ImAddress);
    }
}
