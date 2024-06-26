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
///     Represents an abstract delegate management request.
/// </summary>
/// <typeparam name="TResponse">The type of the response.</typeparam>
internal abstract class DelegateManagementRequestBase<TResponse> : SimpleServiceRequestBase
    where TResponse : DelegateManagementResponse
{
    /// <summary>
    ///     Gets or sets the mailbox.
    /// </summary>
    /// <value>The mailbox.</value>
    public Mailbox Mailbox { get; init; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="DelegateManagementRequestBase&lt;TResponse&gt;" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    internal DelegateManagementRequestBase(ExchangeService service)
        : base(service)
    {
    }

    /// <summary>
    ///     Validate request.
    /// </summary>
    internal override void Validate()
    {
        base.Validate();
        EwsUtilities.ValidateParam(Mailbox);
    }

    /// <summary>
    ///     Writes XML elements.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        Mailbox.WriteToXml(writer, XmlNamespace.Messages, XmlElementNames.Mailbox);
    }

    /// <summary>
    ///     Creates the response.
    /// </summary>
    /// <returns>Response object.</returns>
    internal abstract TResponse CreateResponse();

    /// <summary>
    ///     Parses the response.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>Response object.</returns>
    internal override object ParseResponse(EwsServiceXmlReader reader)
    {
        DelegateManagementResponse response = CreateResponse();

        response.LoadFromXml(reader, GetResponseXmlElementName());

        return response;
    }

    /// <summary>
    ///     Executes this request.
    /// </summary>
    /// <returns>Response object.</returns>
    internal async Task<TResponse> Execute(CancellationToken token)
    {
        var serviceResponse = await InternalExecuteAsync<TResponse>(token).ConfigureAwait(false);

        serviceResponse.ThrowIfNecessary();

        return serviceResponse;
    }
}
