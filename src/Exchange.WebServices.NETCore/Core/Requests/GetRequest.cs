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
///     Represents an abstract Get request.
/// </summary>
/// <typeparam name="TServiceObject">The type of the service object.</typeparam>
/// <typeparam name="TResponse">The type of the response.</typeparam>
internal abstract class GetRequest<TServiceObject, TResponse> : MultiResponseServiceRequest<TResponse>
    where TServiceObject : ServiceObject
    where TResponse : ServiceResponse
{
    /// <summary>
    ///     Gets or sets the property set.
    /// </summary>
    /// <value>The property set.</value>
    public PropertySet PropertySet { get; set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="GetRequest&lt;TServiceObject, TResponse&gt;" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
    internal GetRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
        : base(service, errorHandlingMode)
    {
    }

    /// <summary>
    ///     Validate request.
    /// </summary>
    internal override void Validate()
    {
        base.Validate();
        EwsUtilities.ValidateParam(PropertySet);
        PropertySet.ValidateForRequest(this, false);
    }

    /// <summary>
    ///     Gets the type of the service object this request applies to.
    /// </summary>
    /// <returns>The type of service object the request applies to.</returns>
    internal abstract ServiceObjectType GetServiceObjectType();

    /// <summary>
    ///     Writes XML elements.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        PropertySet.WriteToXml(writer, GetServiceObjectType());
    }
}
