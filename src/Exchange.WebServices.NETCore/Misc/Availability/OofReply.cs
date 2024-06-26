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

using System.Globalization;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents an Out of Office response.
/// </summary>
[PublicAPI]
public sealed class OofReply
{
    /// <summary>
    ///     Gets or sets the culture of the reply.
    /// </summary>
    public string Culture { get; set; } = CultureInfo.CurrentCulture.Name;

    /// <summary>
    ///     Gets or sets the reply message.
    /// </summary>
    public string Message { get; set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="OofReply" /> class.
    /// </summary>
    public OofReply()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="OofReply" /> class.
    /// </summary>
    /// <param name="message">The reply message.</param>
    public OofReply(string message)
    {
        Message = message;
    }

    /// <summary>
    ///     Writes an empty OofReply to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    internal static void WriteEmptyReplyToXml(EwsServiceXmlWriter writer, string xmlElementName)
    {
        writer.WriteStartElement(XmlNamespace.Types, xmlElementName);
        writer.WriteEndElement(); // xmlElementName
    }

    /// <summary>
    ///     Defines an implicit conversion between string an OofReply.
    /// </summary>
    /// <param name="message">The message to convert into OofReply.</param>
    /// <returns>An OofReply initialized with the specified message.</returns>
    public static implicit operator OofReply(string message)
    {
        return new OofReply(message);
    }

    /// <summary>
    ///     Defines an implicit conversion between OofReply and string.
    /// </summary>
    /// <param name="oofReply">The OofReply to convert into a string.</param>
    /// <returns>A string containing the message of the specified OofReply.</returns>
    public static implicit operator string(OofReply oofReply)
    {
        EwsUtilities.ValidateParam(oofReply);

        return oofReply.Message;
    }

    /// <summary>
    ///     Loads from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    internal void LoadFromXml(EwsServiceXmlReader reader, string xmlElementName)
    {
        reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, xmlElementName);

        if (reader.HasAttributes)
        {
            Culture = reader.ReadAttributeValue("xml:lang");
        }

        Message = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Message);

        reader.ReadEndElement(XmlNamespace.Types, xmlElementName);
    }

    /// <summary>
    ///     Writes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    internal void WriteToXml(EwsServiceXmlWriter writer, string xmlElementName)
    {
        writer.WriteStartElement(XmlNamespace.Types, xmlElementName);

        if (Culture != null)
        {
            writer.WriteAttributeValue("xml", "lang", Culture);
        }

        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Message, Message);

        writer.WriteEndElement(); // xmlElementName
    }

    /// <summary>
    ///     Obtains a string representation of the reply.
    /// </summary>
    /// <returns>A string containing the reply message.</returns>
    public override string ToString()
    {
        return Message;
    }
}
