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

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents the Id of a phone call.
/// </summary>
internal sealed class PhoneCallId : ComplexProperty
{
    /// <summary>
    ///     Gets or sets the Id of the phone call.
    /// </summary>
    internal string Id { get; set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="PhoneCallId" /> class.
    /// </summary>
    internal PhoneCallId()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="PhoneCallId" /> class.
    /// </summary>
    /// <param name="id">The Id of the phone call.</param>
    internal PhoneCallId(string id)
    {
        Id = id;
    }

    /// <summary>
    ///     Reads attributes from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
    {
        Id = reader.ReadAttributeValue(XmlAttributeNames.Id);
    }

    /// <summary>
    ///     Writes attributes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteAttributeValue(XmlAttributeNames.Id, Id);
    }

    /// <summary>
    ///     Writes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal void WriteToXml(EwsServiceXmlWriter writer)
    {
        WriteToXml(writer, XmlElementNames.PhoneCallId);
    }
}
