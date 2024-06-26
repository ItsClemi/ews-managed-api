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
///     Represents a GetClientAccessToken request.
/// </summary>
internal sealed class GetClientAccessTokenRequest : MultiResponseServiceRequest<GetClientAccessTokenResponse>
{
    internal ClientAccessTokenRequest[] TokenRequests { get; set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="GetClientAccessTokenRequest" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
    internal GetClientAccessTokenRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
        : base(service, errorHandlingMode)
    {
    }

    /// <summary>
    ///     Creates the service response.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="responseIndex">Index of the response.</param>
    /// <returns>Response object.</returns>
    protected override GetClientAccessTokenResponse CreateServiceResponse(ExchangeService service, int responseIndex)
    {
        return new GetClientAccessTokenResponse(
            TokenRequests[responseIndex].Id,
            TokenRequests[responseIndex].TokenType
        );
    }

    /// <summary>
    ///     Gets the name of the XML element.
    /// </summary>
    /// <returns>XML element name,</returns>
    internal override string GetXmlElementName()
    {
        return XmlElementNames.GetClientAccessToken;
    }

    /// <summary>
    ///     Gets the name of the response XML element.
    /// </summary>
    /// <returns>Xml element name.</returns>
    internal override string GetResponseXmlElementName()
    {
        return XmlElementNames.GetClientAccessTokenResponse;
    }

    /// <summary>
    ///     Gets the name of the response message XML element.
    /// </summary>
    /// <returns>Xml element name.</returns>
    protected override string GetResponseMessageXmlElementName()
    {
        return XmlElementNames.GetClientAccessTokenResponseMessage;
    }

    /// <summary>
    ///     Gets the expected response message count.
    /// </summary>
    /// <returns>Number of items in response.</returns>
    protected override int GetExpectedResponseMessageCount()
    {
        return TokenRequests.Length;
    }

    /// <summary>
    ///     Writes XML elements.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.TokenRequests);

        foreach (var tokenRequestInfo in TokenRequests)
        {
            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.TokenRequest);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Id, tokenRequestInfo.Id);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.TokenType, tokenRequestInfo.TokenType);

            if (!string.IsNullOrEmpty(tokenRequestInfo.Scope))
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.HighlightTermScope,
                    tokenRequestInfo.Scope
                );
            }

            writer.WriteEndElement();
        }

        writer.WriteEndElement();
    }

    /// <summary>
    ///     Validate request.
    /// </summary>
    internal override void Validate()
    {
        base.Validate();

        if (TokenRequests == null || TokenRequests.Length == 0)
        {
            throw new ServiceValidationException(Strings.HoldIdParameterIsNotSpecified);
        }
    }

    /// <summary>
    ///     Gets the request version.
    /// </summary>
    /// <returns>Earliest Exchange version in which this request is supported.</returns>
    internal override ExchangeVersion GetMinimumRequiredServerVersion()
    {
        return ExchangeVersion.Exchange2013;
    }
}
