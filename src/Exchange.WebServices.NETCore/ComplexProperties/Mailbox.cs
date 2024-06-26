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

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a mailbox reference.
/// </summary>
[PublicAPI]
public class Mailbox : ComplexProperty, ISearchStringProvider, IEquatable<Mailbox>
{
    /// <summary>
    ///     True if this instance is valid, false otherwise.
    /// </summary>
    /// <value><c>true</c> if this instance is valid; otherwise, <c>false</c>.</value>
    [MemberNotNullWhen(true, nameof(Address))]
    public bool IsValid => !string.IsNullOrEmpty(Address);

    /// <summary>
    ///     Gets or sets the address used to refer to the user mailbox.
    /// </summary>
    public string? Address { get; set; }

    /// <summary>
    ///     Gets or sets the routing type of the address used to refer to the user mailbox.
    /// </summary>
    public string? RoutingType { get; set; }


    /// <summary>
    ///     Initializes a new instance of the <see cref="Mailbox" /> class.
    /// </summary>
    public Mailbox()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="Mailbox" /> class.
    /// </summary>
    /// <param name="smtpAddress">The primary SMTP address of the mailbox.</param>
    public Mailbox(string? smtpAddress)
        : this()
    {
        Address = smtpAddress;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="Mailbox" /> class.
    /// </summary>
    /// <param name="address">The address used to reference the user mailbox.</param>
    /// <param name="routingType">The routing type of the address used to reference the user mailbox.</param>
    public Mailbox(string address, string routingType)
        : this(address)
    {
        RoutingType = routingType;
    }


    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.EmailAddress:
            {
                Address = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.RoutingType:
            {
                RoutingType = reader.ReadElementValue();
                return true;
            }
            default:
            {
                return false;
            }
        }
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.EmailAddress, Address);
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.RoutingType, RoutingType);
    }


    /// <summary>
    ///     Get a string representation for using this instance in a search filter.
    /// </summary>
    /// <returns>String representation of instance.</returns>
    string ISearchStringProvider.GetSearchString()
    {
        return Address;
    }


    /// <summary>
    ///     Validates this instance.
    /// </summary>
    internal override void InternalValidate()
    {
        base.InternalValidate();

        EwsUtilities.ValidateNonBlankStringParamAllowNull(Address, "address");
        EwsUtilities.ValidateNonBlankStringParamAllowNull(RoutingType, "routingType");
    }


    /// <summary>
    ///     Determines whether the specified <see cref="T:System.Object" /> is equal to the current
    ///     <see cref="T:System.Object" />.
    /// </summary>
    /// <param name="obj">The <see cref="T:System.Object" /> to compare with the current <see cref="T:System.Object" />.</param>
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

        return obj is Mailbox other && Equals(other);
    }

    public bool Equals(Mailbox? other)
    {
        if (other is null)
        {
            return false;
        }

        if (ReferenceEquals(this, other))
        {
            return true;
        }

        if ((Address == null && other.Address == null) || (Address != null && Address.Equals(other.Address)))
        {
            return (RoutingType == null && other.RoutingType == null) ||
                   (RoutingType != null && RoutingType.Equals(other.RoutingType));
        }

        return false;
    }


    /// <summary>
    ///     Serves as a hash function for a particular type.
    /// </summary>
    /// <returns>
    ///     A hash code for the current <see cref="T:System.Object" />.
    /// </returns>
    public override int GetHashCode()
    {
        if (string.IsNullOrEmpty(Address))
        {
            return base.GetHashCode();
        }

        return HashCode.Combine(Address, RoutingType);
    }

    /// <summary>
    ///     Returns a <see cref="T:System.String" /> that represents the current <see cref="T:System.Object" />.
    /// </summary>
    /// <returns>
    ///     A <see cref="T:System.String" /> that represents the current <see cref="T:System.Object" />.
    /// </returns>
    public override string ToString()
    {
        if (!IsValid)
        {
            return string.Empty;
        }

        if (!string.IsNullOrEmpty(RoutingType))
        {
            return RoutingType + ":" + Address;
        }

        return Address;
    }


    /// <summary>
    ///     Defines an implicit conversion between a string representing an SMTP address and Mailbox.
    /// </summary>
    /// <param name="smtpAddress">The SMTP address to convert to EmailAddress.</param>
    /// <returns>A Mailbox initialized with the specified SMTP address.</returns>
    public static implicit operator Mailbox(string smtpAddress)
    {
        return new Mailbox(smtpAddress);
    }

    public static bool operator ==(Mailbox? left, Mailbox? right)
    {
        return Equals(left, right);
    }

    public static bool operator !=(Mailbox? left, Mailbox? right)
    {
        return !Equals(left, right);
    }
}
