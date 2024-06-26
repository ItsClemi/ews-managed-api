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

using Microsoft.Exchange.WebServices.Data;

namespace Microsoft.Exchange.WebServices.Autodiscover;

/// <summary>
///     Represents a GetUserSettings request.
/// </summary>
internal class GetUserSettingsRequest : AutodiscoverRequest
{
    /// <summary>
    ///     Action Uri of Autodiscover.GetUserSettings method.
    /// </summary>
    private const string GetUserSettingsActionUri =
        EwsUtilities.AutodiscoverSoapNamespace + "/Autodiscover/GetUserSettings";

    /// <summary>
    ///     Expect this request to return the partner token.
    /// </summary>
    private readonly bool _expectPartnerToken;

    /// <summary>
    ///     Initializes a new instance of the <see cref="GetUserSettingsRequest" /> class.
    /// </summary>
    /// <param name="service">Autodiscover service associated with this request.</param>
    /// <param name="url">URL of Autodiscover service.</param>
    /// <param name="expectPartnerToken"></param>
    internal GetUserSettingsRequest(AutodiscoverService service, Uri url, bool expectPartnerToken = false)
        : base(service, url)
    {
        _expectPartnerToken = expectPartnerToken;

        // make an explicit https check.
        if (expectPartnerToken && !url.Scheme.Equals("https", StringComparison.OrdinalIgnoreCase))
        {
            throw new ServiceValidationException(Strings.HttpsIsRequired);
        }
    }

    /// <summary>
    ///     Validates the request.
    /// </summary>
    internal override void Validate()
    {
        base.Validate();

        EwsUtilities.ValidateParam(SmtpAddresses);
        EwsUtilities.ValidateParam(Settings);

        if (Settings.Count == 0)
        {
            throw new ServiceValidationException(Strings.InvalidAutodiscoverSettingsCount);
        }

        if (SmtpAddresses.Count == 0)
        {
            throw new ServiceValidationException(Strings.InvalidAutodiscoverSmtpAddressesCount);
        }

        foreach (var smtpAddress in SmtpAddresses)
        {
            if (string.IsNullOrEmpty(smtpAddress))
            {
                throw new ServiceValidationException(Strings.InvalidAutodiscoverSmtpAddress);
            }
        }
    }

    /// <summary>
    ///     Executes this instance.
    /// </summary>
    /// <returns></returns>
    internal async Task<GetUserSettingsResponseCollection> Execute()
    {
        var responses = (GetUserSettingsResponseCollection)await InternalExecute();
        if (responses.ErrorCode == AutodiscoverErrorCode.NoError)
        {
            PostProcessResponses(responses);
        }

        return responses;
    }

    /// <summary>
    ///     Post-process responses to GetUserSettings.
    /// </summary>
    /// <param name="responses">The GetUserSettings responses.</param>
    private void PostProcessResponses(GetUserSettingsResponseCollection responses)
    {
        // Note:The response collection may not include all of the requested users if the request has been throttled.
        for (var index = 0; index < responses.Count; index++)
        {
            responses[index].SmtpAddress = SmtpAddresses[index];
        }
    }

    /// <summary>
    ///     Gets the name of the request XML element.
    /// </summary>
    /// <returns>Request XML element name.</returns>
    internal override string GetRequestXmlElementName()
    {
        return XmlElementNames.GetUserSettingsRequestMessage;
    }

    /// <summary>
    ///     Gets the name of the response XML element.
    /// </summary>
    /// <returns>Response XML element name.</returns>
    internal override string GetResponseXmlElementName()
    {
        return XmlElementNames.GetUserSettingsResponseMessage;
    }

    /// <summary>
    ///     Gets the WS-Addressing action name.
    /// </summary>
    /// <returns>WS-Addressing action name.</returns>
    internal override string GetWsAddressingActionName()
    {
        return GetUserSettingsActionUri;
    }

    /// <summary>
    ///     Creates the service response.
    /// </summary>
    /// <returns>AutodiscoverResponse</returns>
    internal override AutodiscoverResponse CreateServiceResponse()
    {
        return new GetUserSettingsResponseCollection();
    }

    /// <summary>
    ///     Writes the attributes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteAttributeValue(
            "xmlns",
            EwsUtilities.AutodiscoverSoapNamespacePrefix,
            EwsUtilities.AutodiscoverSoapNamespace
        );
    }

    /// <summary>
    /// </summary>
    /// <param name="writer"></param>
    internal override void WriteExtraCustomSoapHeadersToXml(EwsServiceXmlWriter writer)
    {
        if (_expectPartnerToken)
        {
            writer.WriteElementValue(
                XmlNamespace.Autodiscover,
                XmlElementNames.BinarySecret,
                Convert.ToBase64String(ExchangeServiceBase.SessionKey)
            );
        }
    }

    /// <summary>
    ///     Writes request to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteStartElement(XmlNamespace.Autodiscover, XmlElementNames.Request);

        writer.WriteStartElement(XmlNamespace.Autodiscover, XmlElementNames.Users);

        foreach (var smtpAddress in SmtpAddresses)
        {
            writer.WriteStartElement(XmlNamespace.Autodiscover, XmlElementNames.User);

            if (!string.IsNullOrEmpty(smtpAddress))
            {
                writer.WriteElementValue(XmlNamespace.Autodiscover, XmlElementNames.Mailbox, smtpAddress);
            }

            writer.WriteEndElement(); // User
        }

        writer.WriteEndElement(); // Users

        writer.WriteStartElement(XmlNamespace.Autodiscover, XmlElementNames.RequestedSettings);
        foreach (var setting in Settings)
        {
            writer.WriteElementValue(XmlNamespace.Autodiscover, XmlElementNames.Setting, setting);
        }

        writer.WriteEndElement(); // RequestedSettings

        writer.WriteEndElement(); // Request
    }

    /// <summary>
    ///     Read the partner token soap header.
    /// </summary>
    /// <param name="reader">EwsXmlReader</param>
    internal override void ReadSoapHeader(EwsXmlReader reader)
    {
        base.ReadSoapHeader(reader);

        if (_expectPartnerToken)
        {
            if (reader.IsStartElement(XmlNamespace.Autodiscover, XmlElementNames.PartnerToken))
            {
                PartnerToken = reader.ReadInnerXml();
            }

            if (reader.IsStartElement(XmlNamespace.Autodiscover, XmlElementNames.PartnerTokenReference))
            {
                PartnerTokenReference = reader.ReadInnerXml();
            }
        }
    }

    /// <summary>
    ///     Gets or sets the SMTP addresses.
    /// </summary>
    internal List<string> SmtpAddresses { get; set; }

    /// <summary>
    ///     Gets or sets the settings.
    /// </summary>
    internal List<UserSettingName> Settings { get; set; }

    /// <summary>
    ///     Gets the partner token.
    /// </summary>
    internal string PartnerToken { get; private set; }

    /// <summary>
    ///     Gets the partner token reference.
    /// </summary>
    internal string PartnerTokenReference { get; private set; }
}
