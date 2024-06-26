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
///     Represents an impersonated user Id.
/// </summary>
[PublicAPI]
public sealed class ImpersonatedUserId
{
    /// <summary>
    ///     Gets or sets the type of the Id.
    /// </summary>
    public ConnectingIdType IdType { get; set; }

    /// <summary>
    ///     Gets or sets the user Id.
    /// </summary>
    public string Id { get; set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ImpersonatedUserId" /> class.
    /// </summary>
    public ImpersonatedUserId()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ImpersonatedUserId" /> class.
    /// </summary>
    /// <param name="idType">The type of this Id.</param>
    /// <param name="id">The user Id.</param>
    public ImpersonatedUserId(ConnectingIdType idType, string id)
        : this()
    {
        IdType = idType;
        Id = id;
    }

    /// <summary>
    ///     Writes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal void WriteToXml(EwsServiceXmlWriter writer)
    {
        if (string.IsNullOrEmpty(Id))
        {
            throw new ArgumentException(Strings.IdPropertyMustBeSet);
        }

        writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.ExchangeImpersonation);
        writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.ConnectingSID);

        // For 2007 SP1, use PrimarySmtpAddress for type SmtpAddress
        var connectingIdTypeLocalName =
            IdType == ConnectingIdType.SmtpAddress &&
            writer.Service.RequestedServerVersion == ExchangeVersion.Exchange2007_SP1
                ? XmlElementNames.PrimarySmtpAddress : IdType.ToString();

        writer.WriteElementValue(XmlNamespace.Types, connectingIdTypeLocalName, Id);

        writer.WriteEndElement(); // ConnectingSID
        writer.WriteEndElement(); // ExchangeImpersonation
    }
}
