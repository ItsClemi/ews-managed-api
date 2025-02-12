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
///     Represents a service request that can have multiple responses.
/// </summary>
/// <typeparam name="TResponse">The type of the response.</typeparam>
internal abstract class MultiResponseServiceRequest<TResponse> : SimpleServiceRequestBase
    where TResponse : ServiceResponse
{
    /// <summary>
    ///     Gets a value indicating how errors should be handled.
    /// </summary>
    internal ServiceErrorHandling ErrorHandlingMode { get; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="MultiResponseServiceRequest&lt;TResponse&gt;" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
    protected MultiResponseServiceRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
        : base(service)
    {
        ErrorHandlingMode = errorHandlingMode;
    }

    /// <summary>
    ///     Parses the response.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>Service response collection.</returns>
    internal override object ParseResponse(EwsServiceXmlReader reader)
    {
        var serviceResponses = new ServiceResponseCollection<TResponse>();

        reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.ResponseMessages);

        for (var i = 0; i < GetExpectedResponseMessageCount(); i++)
        {
            // Read ahead to see if we've reached the end of the response messages early.
            reader.Read();
            if (reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.ResponseMessages))
            {
                break;
            }

            var response = CreateServiceResponse(reader.Service, i);

            response.LoadFromXml(reader, GetResponseMessageXmlElementName());

            // Add the response to the list after it has been deserialized because the response
            // list updates an overall result as individual responses are added to it.
            serviceResponses.Add(response);
        }

        // If there's a general error in batch processing,
        // the server will return a single response message containing the error
        // (for example, if the SavedItemFolderId is bogus in a batch CreateItem
        // call). In this case, throw a ServiceResponseException. Otherwise this 
        // is an unexpected server error.
        if (serviceResponses.Count < GetExpectedResponseMessageCount())
        {
            if (serviceResponses.Count == 1 && serviceResponses[0].Result == ServiceResult.Error)
            {
                throw new ServiceResponseException(serviceResponses[0]);
            }

            throw new ServiceXmlDeserializationException(
                string.Format(
                    Strings.TooFewServiceReponsesReturned,
                    GetResponseMessageXmlElementName(),
                    GetExpectedResponseMessageCount(),
                    serviceResponses.Count
                )
            );
        }

        reader.ReadEndElementIfNecessary(XmlNamespace.Messages, XmlElementNames.ResponseMessages);

        return serviceResponses;
    }

    /// <summary>
    ///     Creates the service response.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="responseIndex">Index of the response.</param>
    /// <returns>Service response.</returns>
    protected abstract TResponse CreateServiceResponse(ExchangeService service, int responseIndex);

    /// <summary>
    ///     Gets the name of the response message XML element.
    /// </summary>
    /// <returns>XML element name,</returns>
    protected abstract string GetResponseMessageXmlElementName();

    /// <summary>
    ///     Gets the expected response message count.
    /// </summary>
    /// <returns>Number of expected response messages.</returns>
    protected abstract int GetExpectedResponseMessageCount();

    /// <summary>
    ///     Executes this request.
    /// </summary>
    /// <returns>Service response collection.</returns>
    internal async Task<ServiceResponseCollection<TResponse>> ExecuteAsync(CancellationToken cancellationToken)
    {
        var serviceResponses = await InternalExecuteAsync<ServiceResponseCollection<TResponse>>(cancellationToken)
            .ConfigureAwait(false);

        if (ErrorHandlingMode == ServiceErrorHandling.ThrowOnError)
        {
            EwsUtilities.Assert(
                serviceResponses.Count == 1,
                "MultiResponseServiceRequest.Execute",
                "ServiceErrorHandling.ThrowOnError error handling is only valid for singleton request"
            );

            serviceResponses[0].ThrowIfNecessary();
        }

        return serviceResponses;
    }
}
