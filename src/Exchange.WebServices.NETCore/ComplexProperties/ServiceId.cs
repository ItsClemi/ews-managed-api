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

using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents the Id of an Exchange object.
/// </summary>
[PublicAPI]
public abstract class ServiceId : ComplexProperty
{
    /// <summary>
    ///     True if this instance is valid, false otherwise.
    /// </summary>
    /// <value><c>true</c> if this instance is valid; otherwise, <c>false</c>.</value>
    [MemberNotNullWhen(true, nameof(UniqueId))]
    internal virtual bool IsValid => !string.IsNullOrEmpty(UniqueId);

    /// <summary>
    ///     Gets the unique Id of the Exchange object.
    /// </summary>
    public string? UniqueId { get; internal set; }

    /// <summary>
    ///     Gets the change key associated with the Exchange object. The change key represents the
    ///     the version of the associated item or folder.
    /// </summary>
    public string? ChangeKey { get; internal set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ServiceId" /> class.
    /// </summary>
    internal ServiceId()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ServiceId" /> class.
    /// </summary>
    /// <param name="uniqueId">The unique id.</param>
    internal ServiceId(string uniqueId)
        : this()
    {
        ArgumentException.ThrowIfNullOrEmpty(uniqueId);

        UniqueId = uniqueId;
    }

    /// <summary>
    ///     Reads attributes from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
    {
        UniqueId = reader.ReadAttributeValue(XmlAttributeNames.Id);
        ChangeKey = reader.ReadAttributeValue(XmlAttributeNames.ChangeKey);
    }

    /// <summary>
    ///     Writes attributes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteAttributeValue(XmlAttributeNames.Id, UniqueId);
        writer.WriteAttributeValue(XmlAttributeNames.ChangeKey, ChangeKey);
    }

    /// <summary>
    ///     Gets the name of the XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal abstract string GetXmlElementName();

    /// <summary>
    ///     Writes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal void WriteToXml(EwsServiceXmlWriter writer)
    {
        WriteToXml(writer, GetXmlElementName());
    }

    /// <summary>
    ///     Assigns from existing id.
    /// </summary>
    /// <param name="source">The source.</param>
    internal void Assign(ServiceId source)
    {
        UniqueId = source.UniqueId;
        ChangeKey = source.ChangeKey;
    }

    /// <summary>
    ///     Determines whether two ServiceId instances are equal (including ChangeKeys)
    /// </summary>
    /// <param name="other">The ServiceId to compare with the current ServiceId.</param>
    public bool SameIdAndChangeKey(ServiceId other)
    {
        if (Equals(other))
        {
            return (ChangeKey == null && other.ChangeKey == null) || ChangeKey.Equals(other.ChangeKey);
        }

        return false;
    }


    /// <summary>
    ///     Determines whether the specified <see cref="T:System.Object" /> is equal to the current
    ///     <see cref="T:System.Object" />.
    /// </summary>
    /// <param name="obj">The <see cref="T:System.Object" /> to compare with the current <see cref="T:System.Object" />.</param>
    /// <remarks>
    ///     We do not consider the ChangeKey for ServiceId.Equals.
    /// </remarks>
    /// <returns>
    ///     true if the specified <see cref="T:System.Object" /> is equal to the current <see cref="T:System.Object" />;
    ///     otherwise, false.
    /// </returns>
    /// <exception cref="T:System.NullReferenceException">The <paramref name="obj" /> parameter is null.</exception>
    public override bool Equals(object? obj)
    {
        if (ReferenceEquals(this, obj))
        {
            return true;
        }

        if (obj is not ServiceId other)
        {
            return false;
        }

        if (!(IsValid && other.IsValid))
        {
            return false;
        }

        return UniqueId.Equals(other.UniqueId);
    }

    /// <summary>
    ///     Serves as a hash function for a particular type.
    /// </summary>
    /// <remarks>
    ///     We do not consider the change key in the hash code computation.
    /// </remarks>
    /// <returns>
    ///     A hash code for the current <see cref="T:System.Object" />.
    /// </returns>
    public override int GetHashCode()
    {
        if (IsValid)
        {
            return UniqueId.GetHashCode();
        }

        return RuntimeHelpers.GetHashCode(this);
    }

    /// <summary>
    ///     Returns a <see cref="T:System.String" /> that represents the current <see cref="T:System.Object" />.
    /// </summary>
    /// <returns>
    ///     A <see cref="T:System.String" /> that represents the current <see cref="T:System.Object" />.
    /// </returns>
    public override string ToString()
    {
        return UniqueId ?? string.Empty;
    }
}
