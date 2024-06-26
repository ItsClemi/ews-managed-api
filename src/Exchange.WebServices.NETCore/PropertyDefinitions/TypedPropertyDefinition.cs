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
///     Represents typed property definition.
/// </summary>
internal abstract class TypedPropertyDefinition : PropertyDefinition
{
    /// <summary>
    ///     Gets a value indicating whether this property definition is for a nullable type (ref, int?, bool?...).
    /// </summary>
    internal override bool IsNullable { get; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="TypedPropertyDefinition" /> class.
    /// </summary>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <param name="uri">The URI.</param>
    /// <param name="version">The version.</param>
    internal TypedPropertyDefinition(string xmlElementName, string uri, ExchangeVersion version)
        : base(xmlElementName, uri, version)
    {
        IsNullable = false;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="TypedPropertyDefinition" /> class.
    /// </summary>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <param name="uri">The URI.</param>
    /// <param name="flags">The flags.</param>
    /// <param name="version">The version.</param>
    internal TypedPropertyDefinition(
        string xmlElementName,
        string uri,
        PropertyDefinitionFlags flags,
        ExchangeVersion version
    )
        : base(xmlElementName, uri, flags, version)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="TypedPropertyDefinition" /> class.
    /// </summary>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <param name="uri">The URI.</param>
    /// <param name="flags">The flags.</param>
    /// <param name="version">The version.</param>
    /// <param name="isNullable">Indicates that this property definition is for a nullable property.</param>
    internal TypedPropertyDefinition(
        string xmlElementName,
        string uri,
        PropertyDefinitionFlags flags,
        ExchangeVersion version,
        bool isNullable
    )
        : this(xmlElementName, uri, flags, version)
    {
        IsNullable = isNullable;
    }

    /// <summary>
    ///     Parses the specified value.
    /// </summary>
    /// <param name="value">The value.</param>
    /// <returns>Typed value.</returns>
    internal abstract object Parse(string value);

    /// <summary>
    ///     Convert instance to string.
    /// </summary>
    /// <param name="value">The value.</param>
    /// <returns>String representation of property value.</returns>
    internal virtual string ToString(object value)
    {
        return value.ToString();
    }

    /// <summary>
    ///     Loads from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="propertyBag">The property bag.</param>
    internal override void LoadPropertyValueFromXml(EwsServiceXmlReader reader, PropertyBag propertyBag)
    {
        var value = reader.ReadElementValue(XmlNamespace.Types, XmlElementName);

        if (!string.IsNullOrEmpty(value))
        {
            propertyBag[this] = Parse(value);
        }
    }

    /// <summary>
    ///     Writes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="propertyBag">The property bag.</param>
    /// <param name="isUpdateOperation">Indicates whether the context is an update operation.</param>
    internal override void WritePropertyValueToXml(
        EwsServiceXmlWriter writer,
        PropertyBag propertyBag,
        bool isUpdateOperation
    )
    {
        var value = propertyBag[this];

        if (value != null)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementName, Name, value);
        }
    }
}
