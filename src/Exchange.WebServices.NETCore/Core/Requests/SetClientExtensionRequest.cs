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
///     Represents a SetClientExtension request.
/// </summary>
internal sealed class SetClientExtensionRequest : MultiResponseServiceRequest<ServiceResponse>
{
    /// <summary>
    ///     Set action such as install, uninstall and configure.
    /// </summary>
    private readonly List<SetClientExtensionAction> _actions;

    /// <summary>
    ///     Initializes a new instance of the <see cref="SetClientExtensionRequest" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="actions">List of actions to execute.</param>
    internal SetClientExtensionRequest(ExchangeService service, List<SetClientExtensionAction> actions)
        : base(service, ServiceErrorHandling.ThrowOnError)
    {
        _actions = actions;
    }

    /// <summary>
    ///     Validate request.
    /// </summary>
    internal override void Validate()
    {
        base.Validate();
        EwsUtilities.ValidateParam(_actions);
    }

    /// <summary>
    ///     Creates the service response.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="responseIndex">Index of the response.</param>
    /// <returns>Service response.</returns>
    protected override ServiceResponse CreateServiceResponse(ExchangeService service, int responseIndex)
    {
        return new ServiceResponse();
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
    ///     Gets the expected response message count.
    /// </summary>
    /// <returns>Number of expected response messages.</returns>
    protected override int GetExpectedResponseMessageCount()
    {
        return _actions.Count;
    }

    /// <summary>
    ///     Gets the name of the XML element.
    /// </summary>
    /// <returns>XML element name,</returns>
    internal override string GetXmlElementName()
    {
        return XmlElementNames.SetClientExtensionRequest;
    }

    /// <summary>
    ///     Gets the name of the response XML element.
    /// </summary>
    /// <returns>XML element name,</returns>
    internal override string GetResponseXmlElementName()
    {
        return XmlElementNames.SetClientExtensionResponse;
    }

    /// <summary>
    ///     Gets the name of the response message XML element.
    /// </summary>
    /// <returns>XML element name,</returns>
    protected override string GetResponseMessageXmlElementName()
    {
        return XmlElementNames.SetClientExtensionResponseMessage;
    }

    /// <summary>
    ///     Writes XML elements.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.SetClientExtensionActions);

        foreach (var action in _actions)
        {
            action.WriteToXml(writer, XmlElementNames.SetClientExtensionAction);
        }

        writer.WriteEndElement();
    }
}
