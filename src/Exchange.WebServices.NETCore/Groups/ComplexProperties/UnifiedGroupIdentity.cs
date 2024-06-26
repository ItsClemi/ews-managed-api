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

namespace Microsoft.Exchange.WebServices.Data.Groups;

/// <summary>
///     Defines the UnifiedGroupIdentity class.
/// </summary>
internal sealed class UnifiedGroupIdentity : ComplexProperty
{
    /// <summary>
    ///     Gets or sets the IdentityType of the UnifiedGroup
    /// </summary>
    public UnifiedGroupIdentityType IdentityType { get; set; }

    /// <summary>
    ///     Gets or sets the value associated with the IdentityType for the UnifiedGroup
    /// </summary>
    public string Value { get; set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="UnifiedGroupIdentity" />  class
    /// </summary>
    /// <param name="identityType">The identity type</param>
    /// <param name="value">The value associated with the identity type</param>
    public UnifiedGroupIdentity(UnifiedGroupIdentityType identityType, string value)
    {
        IdentityType = identityType;
        Value = value;
    }

    /// <summary>
    ///     Writes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    internal override void WriteToXml(EwsServiceXmlWriter writer, string xmlElementName)
    {
        writer.WriteStartElement(XmlNamespace.Types, xmlElementName);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.GroupIdentityType, IdentityType.ToString());
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.GroupIdentityValue, Value);
        writer.WriteEndElement();
    }
}
