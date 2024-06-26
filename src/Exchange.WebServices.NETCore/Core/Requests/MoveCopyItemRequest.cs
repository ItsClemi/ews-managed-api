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
///     Represents an abstract Move/Copy Item request.
/// </summary>
/// <typeparam name="TResponse">The type of the response.</typeparam>
internal abstract class MoveCopyItemRequest<TResponse> : MoveCopyRequest<Item, TResponse>
    where TResponse : ServiceResponse
{
    /// <summary>
    ///     Gets the item ids.
    /// </summary>
    /// <value>The item ids.</value>
    internal ItemIdWrapperList ItemIds { get; } = new();

    /// <summary>
    ///     Gets or sets flag indicating whether we require that the service return new item ids.
    /// </summary>
    internal bool? ReturnNewItemIds { get; set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="MoveCopyItemRequest&lt;TResponse&gt;" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
    internal MoveCopyItemRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
        : base(service, errorHandlingMode)
    {
    }

    /// <summary>
    ///     Validates request.
    /// </summary>
    internal override void Validate()
    {
        base.Validate();

        EwsUtilities.ValidateParam(ItemIds);
    }

    /// <summary>
    ///     Writes the ids and returnNewItemIds flag as XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteIdsToXml(EwsServiceXmlWriter writer)
    {
        ItemIds.WriteToXml(writer, XmlNamespace.Messages, XmlElementNames.ItemIds);

        if (ReturnNewItemIds.HasValue)
        {
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.ReturnNewItemIds, ReturnNewItemIds.Value);
        }
    }

    /// <summary>
    ///     Gets the expected response message count.
    /// </summary>
    /// <returns>Number of expected response messages.</returns>
    protected override int GetExpectedResponseMessageCount()
    {
        return ItemIds.Count;
    }
}
