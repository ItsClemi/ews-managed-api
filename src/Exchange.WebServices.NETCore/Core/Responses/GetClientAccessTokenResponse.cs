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
///     Represents the response to a GetClientAccessToken operation.
/// </summary>
[PublicAPI]
public sealed class GetClientAccessTokenResponse : ServiceResponse
{
    /// <summary>
    ///     Gets the Id.
    /// </summary>
    public string Id { get; private set; }

    /// <summary>
    ///     Gets the token type.
    /// </summary>
    public ClientAccessTokenType TokenType { get; private set; }

    /// <summary>
    ///     Gets the token value.
    /// </summary>
    public string TokenValue { get; private set; }

    /// <summary>
    ///     Gets the TTL value in minutes.
    /// </summary>
    public int TTL { get; private set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="GetClientAccessTokenResponse" /> class.
    /// </summary>
    /// <param name="id">Id</param>
    /// <param name="tokenType">Token type</param>
    internal GetClientAccessTokenResponse(string id, ClientAccessTokenType tokenType)
    {
        Id = id;
        TokenType = tokenType;
    }

    /// <summary>
    ///     Reads response elements from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
    {
        base.ReadElementsFromXml(reader);

        reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.Token);

        Id = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Id);
        TokenType = Enum.Parse<ClientAccessTokenType>(
            reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.TokenType)
        );
        TokenValue = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.TokenValue);
        TTL = int.Parse(reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.TTL));

        reader.ReadEndElementIfNecessary(XmlNamespace.Messages, XmlElementNames.Token);
    }
}
