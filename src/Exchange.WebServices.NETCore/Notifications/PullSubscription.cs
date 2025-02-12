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
///     Represents a pull subscription.
/// </summary>
[PublicAPI]
public sealed class PullSubscription : SubscriptionBase
{
    /// <summary>
    ///     Gets a value indicating whether more events are available on the server.
    ///     MoreEventsAvailable is undefined (null) until GetEvents is called.
    /// </summary>
    public bool? MoreEventsAvailable { get; private set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="PullSubscription" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    internal PullSubscription(ExchangeService service)
        : base(service)
    {
    }

    /// <summary>
    ///     Obtains a collection of events that occurred on the subscribed folders since the point
    ///     in time defined by the Watermark property. When GetEvents succeeds, Watermark is updated.
    /// </summary>
    /// <returns>Returns a collection of events that occurred since the last watermark.</returns>
    public async Task<GetEventsResults> GetEvents(CancellationToken token = default)
    {
        var results = await Service.GetEvents(Id, Watermark, token).ConfigureAwait(false);

        Watermark = results.NewWatermark;
        MoreEventsAvailable = results.MoreEventsAvailable;

        return results;
    }

    /// <summary>
    ///     Unsubscribes from the pull subscription.
    /// </summary>
    public System.Threading.Tasks.Task Unsubscribe(CancellationToken token = default)
    {
        return Service.Unsubscribe(Id, token);
    }
}
