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
///     Represents a FindFolder request.
/// </summary>
internal sealed class FindFolderRequest : FindRequest<FindFolderResponse>
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="FindFolderRequest" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
    internal FindFolderRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
        : base(service, errorHandlingMode)
    {
    }

    /// <summary>
    ///     Creates the service response.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="responseIndex">Index of the response.</param>
    /// <returns>Service response.</returns>
    protected override FindFolderResponse CreateServiceResponse(ExchangeService service, int responseIndex)
    {
        return new FindFolderResponse(View.GetPropertySetOrDefault());
    }

    /// <summary>
    ///     Gets the name of the XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal override string GetXmlElementName()
    {
        return XmlElementNames.FindFolder;
    }

    /// <summary>
    ///     Gets the name of the response XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal override string GetResponseXmlElementName()
    {
        return XmlElementNames.FindFolderResponse;
    }

    /// <summary>
    ///     Gets the name of the response message XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    protected override string GetResponseMessageXmlElementName()
    {
        return XmlElementNames.FindFolderResponseMessage;
    }

    /// <summary>
    ///     Gets the request version.
    /// </summary>
    /// <returns>Earliest Exchange version in which this request is supported.</returns>
    internal override ExchangeVersion GetMinimumRequiredServerVersion()
    {
        return ExchangeVersion.Exchange2007_SP1;
    }
}
