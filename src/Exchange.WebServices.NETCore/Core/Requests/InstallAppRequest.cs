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
///     Represents a InstallApp request.
/// </summary>
internal sealed class InstallAppRequest : SimpleServiceRequestBase
{
    /// <summary>
    ///     The plain text manifest stream.
    /// </summary>
    private readonly Stream _manifestStream;

    /// <summary>
    ///     The asset id of the addin in marketplace
    /// </summary>
    private readonly string _marketplaceAssetId;

    /// <summary>
    ///     The target market for content
    /// </summary>
    private readonly string _marketplaceContentMarket;

    /// <summary>
    ///     Whether to send welcome email or not
    /// </summary>
    private readonly bool _sendWelcomeEmail;

    /// <summary>
    ///     Initializes a new instance of the <see cref="InstallAppRequest" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="manifestStream">The manifest's plain text XML stream. </param>
    /// <param name="marketplaceAssetId">The asset id of the addin in marketpalce</param>
    /// <param name="marketplaceContentMarket">The target market for the content</param>
    /// <param name="sendWelcomeEmail">Whether to send email on installation</param>
    internal InstallAppRequest(
        ExchangeService service,
        Stream manifestStream,
        string marketplaceAssetId,
        string marketplaceContentMarket,
        bool sendWelcomeEmail
    )
        : base(service)
    {
        _manifestStream = manifestStream;
        _marketplaceAssetId = marketplaceAssetId;
        _marketplaceContentMarket = marketplaceContentMarket;
        _sendWelcomeEmail = sendWelcomeEmail;
    }

    /// <summary>
    ///     Gets the name of the XML element.
    /// </summary>
    /// <returns>XML element name,</returns>
    internal override string GetXmlElementName()
    {
        return XmlElementNames.InstallAppRequest;
    }

    /// <summary>
    ///     Writes XML elements.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.Manifest);

        writer.WriteBase64ElementValue(_manifestStream);

        if (!string.IsNullOrEmpty(_marketplaceAssetId))
        {
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.MarketplaceAssetId, _marketplaceAssetId);

            if (!string.IsNullOrEmpty(_marketplaceContentMarket))
            {
                writer.WriteElementValue(
                    XmlNamespace.Messages,
                    XmlElementNames.MarketplaceContentMarket,
                    _marketplaceContentMarket
                );
            }

            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.SendWelcomeEmail, _sendWelcomeEmail);
        }

        writer.WriteEndElement();
    }

    /// <summary>
    ///     Gets the name of the response XML element.
    /// </summary>
    /// <returns>XML element name,</returns>
    internal override string GetResponseXmlElementName()
    {
        return XmlElementNames.InstallAppResponse;
    }

    /// <summary>
    ///     Parses the response.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>Response object.</returns>
    internal override object ParseResponse(EwsServiceXmlReader reader)
    {
        var response = new InstallAppResponse();
        response.LoadFromXml(reader, XmlElementNames.InstallAppResponse);
        return response;
    }

    /// <summary>
    ///     Gets the request version.
    /// </summary>
    /// <returns>Earliest Exchange version in which this request is supported.</returns>
    internal override ExchangeVersion GetMinimumRequiredServerVersion()
    {
        return ExchangeVersion.Exchange2013;
    }

    /// <summary>
    ///     Executes this request.
    /// </summary>
    /// <returns>Service response.</returns>
    internal async Task<InstallAppResponse> Execute(CancellationToken token)
    {
        var serviceResponse = await InternalExecuteAsync<InstallAppResponse>(token).ConfigureAwait(false);
        serviceResponse.ThrowIfNecessary();
        return serviceResponse;
    }
}
