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
///     Represents an indexed property definition.
/// </summary>
[PublicAPI]
public sealed class IndexedPropertyDefinition : ServiceObjectPropertyDefinition
{
    /// <summary>
    ///     Gets the index attribute of IndexedFieldURI element.
    /// </summary>
    public string Index { get; }

    /// <summary>
    ///     Gets the property type.
    /// </summary>
    public override Type Type => typeof(string);

    /// <summary>
    ///     Initializes a new instance of the <see cref="IndexedPropertyDefinition" /> class.
    /// </summary>
    /// <param name="uri">The FieldURI attribute of the IndexedFieldURI element.</param>
    /// <param name="index">The Index attribute of the IndexedFieldURI element.</param>
    internal IndexedPropertyDefinition(string uri, string index)
        : base(uri)
    {
        Index = index;
    }

    /// <summary>
    ///     Determines whether two specified instances of IndexedPropertyDefinition are equal.
    /// </summary>
    /// <param name="idxPropDef1">First indexed property definition.</param>
    /// <param name="idxPropDef2">Second indexed property definition.</param>
    /// <returns>True if indexed property definitions are equal.</returns>
    internal static bool IsEqualTo(IndexedPropertyDefinition? idxPropDef1, IndexedPropertyDefinition? idxPropDef2)
    {
        return ReferenceEquals(idxPropDef1, idxPropDef2) ||
               (idxPropDef1 is not null &&
                idxPropDef2 is not null &&
                idxPropDef1.Uri == idxPropDef2.Uri &&
                idxPropDef1.Index == idxPropDef2.Index);
    }

    /// <summary>
    ///     Writes the attributes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        base.WriteAttributesToXml(writer);

        writer.WriteAttributeValue(XmlAttributeNames.FieldIndex, Index);
    }

    /// <summary>
    ///     Gets the name of the XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal override string GetXmlElementName()
    {
        return XmlElementNames.IndexedFieldURI;
    }

    /// <summary>
    ///     Gets the property definition's printable name.
    /// </summary>
    /// <returns>
    ///     The property definition's printable name.
    /// </returns>
    internal override string GetPrintableName()
    {
        return $"{Uri}:{Index}";
    }

    /// <summary>
    ///     Determines whether two specified instances of IndexedPropertyDefinition are equal.
    /// </summary>
    /// <param name="idxPropDef1">First indexed property definition.</param>
    /// <param name="idxPropDef2">Second indexed property definition.</param>
    /// <returns>True if indexed property definitions are equal.</returns>
    public static bool operator ==(IndexedPropertyDefinition idxPropDef1, IndexedPropertyDefinition idxPropDef2)
    {
        return IsEqualTo(idxPropDef1, idxPropDef2);
    }

    /// <summary>
    ///     Determines whether two specified instances of IndexedPropertyDefinition are not equal.
    /// </summary>
    /// <param name="idxPropDef1">First indexed property definition.</param>
    /// <param name="idxPropDef2">Second indexed property definition.</param>
    /// <returns>True if indexed property definitions are equal.</returns>
    public static bool operator !=(IndexedPropertyDefinition idxPropDef1, IndexedPropertyDefinition idxPropDef2)
    {
        return !IsEqualTo(idxPropDef1, idxPropDef2);
    }

    /// <summary>
    ///     Determines whether a given indexed property definition is equal to this indexed property definition.
    /// </summary>
    /// <param name="obj">The object to check for equality.</param>
    /// <returns>True if the properties definitions define the same indexed property.</returns>
    public override bool Equals(object? obj)
    {
        var propertyDefinition = obj as IndexedPropertyDefinition;
        return IsEqualTo(propertyDefinition, this);
    }

    /// <summary>
    ///     Serves as a hash function for a particular type.
    /// </summary>
    /// <returns>
    ///     A hash code for the current <see cref="T:System.Object" />.
    /// </returns>
    public override int GetHashCode()
    {
        return Uri.GetHashCode() ^ Index.GetHashCode();
    }
}
