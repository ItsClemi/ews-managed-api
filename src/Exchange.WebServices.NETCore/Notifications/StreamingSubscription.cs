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

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a streaming subscription.
/// </summary>
[PublicAPI]
public sealed class StreamingSubscription : SubscriptionBase
{
    /// <summary>
    ///     Gets the service used to create this subscription.
    /// </summary>
    public new ExchangeService Service => base.Service;

    /// <summary>
    ///     Gets a value indicating whether this subscription uses watermarks.
    /// </summary>
    protected override bool UsesWatermark => false;

    /// <summary>
    ///     Initializes a new instance of the <see cref="StreamingSubscription" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    internal StreamingSubscription(ExchangeService service)
        : base(service)
    {
    }

    /// <summary>
    ///     Initializes a new instance with the specified subscription id, for continuing an existing subscription.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="subscriptionId">The id of a previously created streaming subscription.</param>
    public StreamingSubscription(ExchangeService service, string subscriptionId)
        : base(service)
    {
        Id = subscriptionId;
    }

    /// <summary>
    ///     Unsubscribes from the streaming subscription.
    /// </summary>
    public System.Threading.Tasks.Task Unsubscribe(CancellationToken token = default)
    {
        return Service.Unsubscribe(Id, token);
    }
}
