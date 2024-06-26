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

using System.Net;
using System.Net.Http.Headers;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a request of a get user photo operation
/// </summary>
internal sealed class GetUserPhotoRequest : SimpleServiceRequestBase
{
    /// <summary>
    ///     email address accessor
    /// </summary>
    internal string EmailAddress { get; set; }

    /// <summary>
    ///     user photo size accessor
    /// </summary>
    internal string UserPhotoSize { get; set; }

    /// <summary>
    ///     EntityTag accessor
    /// </summary>
    internal string EntityTag { get; set; }

    /// <summary>
    ///     Default constructor
    /// </summary>
    /// <param name="service">Exchange web service</param>
    internal GetUserPhotoRequest(ExchangeService service)
        : base(service)
    {
    }

    /// <summary>
    ///     Creates a NotFound instance of the result
    /// </summary>
    /// <returns>The canonical NotFound result</returns>
    internal static GetUserPhotoResponse GetNotFoundResponse()
    {
        var serviceResponse = new GetUserPhotoResponse
        {
            Results =
            {
                Status = GetUserPhotoStatus.PhotoOrUserNotFound,
            },
        };

        return serviceResponse;
    }

    /// <summary>
    ///     Validate request.
    /// </summary>
    internal override void Validate()
    {
        if (string.IsNullOrEmpty(EmailAddress))
        {
            throw new ServiceLocalException(Strings.InvalidEmailAddress);
        }

        if (string.IsNullOrEmpty(UserPhotoSize))
        {
            throw new ServiceLocalException(Strings.UserPhotoSizeNotSpecified);
        }

        base.Validate();
    }

    /// <summary>
    ///     Writes XML elements.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        // Emit the EmailAddress element
        writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.Email);
        writer.WriteValue(EmailAddress, XmlElementNames.Email);
        writer.WriteEndElement();

        writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.SizeRequested);
        writer.WriteValue(UserPhotoSize, XmlElementNames.SizeRequested);
        writer.WriteEndElement();
    }

    /// <summary>
    ///     Adds header values to the request
    /// </summary>
    /// <param name="webHeaderCollection">The collection of headers to add to</param>
    internal override void AddHeaders(HttpRequestHeaders webHeaderCollection)
    {
        // Check if the ETag was specified
        if (!string.IsNullOrEmpty(EntityTag))
        {
            // Ensure the ETag is wrapped in quotes
            var quotedETag = EntityTag;
            if (!EntityTag.StartsWith("\""))
            {
                quotedETag = "\"" + quotedETag;
            }

            if (!EntityTag.EndsWith("\""))
            {
                quotedETag += "\"";
            }

            webHeaderCollection.IfNoneMatch.ParseAdd(quotedETag);
        }
    }

    /// <summary>
    ///     Parses the response.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="responseHeaders">The HTTP response headers</param>
    /// <returns>Response object.</returns>
    internal override object ParseResponse(EwsServiceXmlReader reader, HttpResponseHeaders responseHeaders)
    {
        var response = new GetUserPhotoResponse();
        response.LoadFromXml(reader, XmlElementNames.GetUserPhotoResponse);
        response.ReadHeader(responseHeaders);
        return response;
    }

    /// <summary>
    ///     Gets the name of the XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal override string GetXmlElementName()
    {
        return XmlElementNames.GetUserPhoto;
    }

    /// <summary>
    ///     Gets the name of the response XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal override string GetResponseXmlElementName()
    {
        return XmlElementNames.GetUserPhotoResponse;
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
    internal async Task<GetUserPhotoResponse> Execute(CancellationToken token)
    {
        try
        {
            return await InternalExecuteAsync<GetUserPhotoResponse>(token).ConfigureAwait(false);
        }
        catch (ServiceRequestException ex)
        {
            // 404 is a valid return code in the case of GetUserPhoto when the photo is
            // not found, so it is necessary to catch this exception here.
            if (ex.InnerException is EwsHttpClientException webException)
            {
                var errorResponse = webException.Response;
                if (errorResponse != null && errorResponse.StatusCode == HttpStatusCode.NotFound)
                {
                    return GetNotFoundResponse();
                }
            }

            throw;
        }
    }
}
