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
///     Represents the base class for Id expressed in a specific format.
/// </summary>
[PublicAPI]
public abstract class AlternateIdBase : ISelfValidate
{
    /// <summary>
    ///     Gets or sets the format in which the Id in expressed.
    /// </summary>
    public IdFormat Format { get; set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="AlternateIdBase" /> class.
    /// </summary>
    internal AlternateIdBase()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="AlternateIdBase" /> class.
    /// </summary>
    /// <param name="format">The format.</param>
    internal AlternateIdBase(IdFormat format)
        : this()
    {
        Format = format;
    }


    #region ISelfValidate Members

    /// <summary>
    ///     Validates this instance.
    /// </summary>
    void ISelfValidate.Validate()
    {
        InternalValidate();
    }

    #endregion


    /// <summary>
    ///     Gets the name of the XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal abstract string GetXmlElementName();

    /// <summary>
    ///     Writes the attributes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal virtual void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteAttributeValue(XmlAttributeNames.Format, Format);
    }

    /// <summary>
    ///     Loads the attributes from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal virtual void LoadAttributesFromXml(EwsServiceXmlReader reader)
    {
        Format = reader.ReadAttributeValue<IdFormat>(XmlAttributeNames.Format);
    }

    /// <summary>
    ///     Writes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal void WriteToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteStartElement(XmlNamespace.Types, GetXmlElementName());

        WriteAttributesToXml(writer);

        writer.WriteEndElement();
    }

    /// <summary>
    ///     Validate this instance.
    /// </summary>
    internal virtual void InternalValidate()
    {
        // nothing to do.
    }
}
