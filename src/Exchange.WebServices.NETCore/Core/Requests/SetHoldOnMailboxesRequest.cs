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
///     Represents a SetHoldOnMailboxesRequest request.
/// </summary>
internal sealed class SetHoldOnMailboxesRequest : SimpleServiceRequestBase
{
    /// <summary>
    ///     Action type
    /// </summary>
    public HoldAction ActionType { get; set; }

    /// <summary>
    ///     Hold id
    /// </summary>
    public string HoldId { get; set; }

    /// <summary>
    ///     Query
    /// </summary>
    public string? Query { get; set; }

    /// <summary>
    ///     Collection of mailboxes to be held/unheld
    /// </summary>
    public string[]? Mailboxes { get; set; }

    /// <summary>
    ///     Query language
    /// </summary>
    public string? Language { get; set; }

    /// <summary>
    ///     InPlaceHold Identity
    /// </summary>
    public string InPlaceHoldIdentity { get; set; }

    /// <summary>
    ///     Item hold period
    /// </summary>
    public string? ItemHoldPeriod { get; set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="SetHoldOnMailboxesRequest" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    internal SetHoldOnMailboxesRequest(ExchangeService service)
        : base(service)
    {
    }

    /// <summary>
    ///     Gets the name of the response XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal override string GetResponseXmlElementName()
    {
        return XmlElementNames.SetHoldOnMailboxesResponse;
    }

    /// <summary>
    ///     Gets the name of the XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal override string GetXmlElementName()
    {
        return XmlElementNames.SetHoldOnMailboxes;
    }

    /// <summary>
    ///     Validate request.
    /// </summary>
    internal override void Validate()
    {
        base.Validate();

        if (string.IsNullOrEmpty(HoldId))
        {
            throw new ServiceValidationException(Strings.HoldIdParameterIsNotSpecified);
        }

        if (string.IsNullOrEmpty(InPlaceHoldIdentity) && (Mailboxes == null || Mailboxes.Length == 0))
        {
            throw new ServiceValidationException(Strings.HoldMailboxesParameterIsNotSpecified);
        }
    }

    /// <summary>
    ///     Parses the response.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>Response object.</returns>
    internal override object ParseResponse(EwsServiceXmlReader reader)
    {
        var response = new SetHoldOnMailboxesResponse();
        response.LoadFromXml(reader, GetResponseXmlElementName());
        return response;
    }

    /// <summary>
    ///     Writes XML elements.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.ActionType, ActionType.ToString());
        writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.HoldId, HoldId);
        writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.Query, Query ?? string.Empty);

        if (Mailboxes != null && Mailboxes.Length > 0)
        {
            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.Mailboxes);
            foreach (var mailbox in Mailboxes)
            {
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.String, mailbox);
            }

            writer.WriteEndElement(); // Mailboxes
        }

        // Language
        if (!string.IsNullOrEmpty(Language))
        {
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.Language, Language);
        }

        if (!string.IsNullOrEmpty(InPlaceHoldIdentity))
        {
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.InPlaceHoldIdentity, InPlaceHoldIdentity);
        }

        if (!string.IsNullOrEmpty(ItemHoldPeriod))
        {
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.ItemHoldPeriod, ItemHoldPeriod);
        }
    }

    /// <summary>
    ///     Executes this request.
    /// </summary>
    /// <returns>Service response.</returns>
    internal async Task<SetHoldOnMailboxesResponse> Execute(CancellationToken token)
    {
        var serviceResponse = await InternalExecuteAsync<SetHoldOnMailboxesResponse>(token).ConfigureAwait(false);
        return serviceResponse;
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
