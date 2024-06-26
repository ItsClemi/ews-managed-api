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
///     Represents the Unified Messaging functionalities.
/// </summary>
[PublicAPI]
public sealed class UnifiedMessaging
{
    private readonly ExchangeService _service;

    /// <summary>
    ///     Constructor
    /// </summary>
    /// <param name="service">EWS service to which this object belongs.</param>
    internal UnifiedMessaging(ExchangeService service)
    {
        _service = service;
    }

    /// <summary>
    ///     Calls a phone and reads a message to the person who picks up.
    /// </summary>
    /// <param name="itemId">The Id of the message to read.</param>
    /// <param name="dialString">The full dial string used to call the phone.</param>
    /// <param name="token"></param>
    /// <returns>An object providing status for the phone call.</returns>
    public async Task<PhoneCall> PlayOnPhone(ItemId itemId, string dialString, CancellationToken token = default)
    {
        EwsUtilities.ValidateParam(itemId);
        EwsUtilities.ValidateParam(dialString);

        var request = new PlayOnPhoneRequest(_service)
        {
            DialString = dialString,
            ItemId = itemId,
        };

        var serviceResponse = await request.Execute(token).ConfigureAwait(false);

        var callInformation = new PhoneCall(_service, serviceResponse.PhoneCallId);

        return callInformation;
    }

    /// <summary>
    ///     Retrieves information about a current phone call.
    /// </summary>
    /// <param name="id">The Id of the phone call.</param>
    /// <param name="token"></param>
    /// <returns>An object providing status for the phone call.</returns>
    internal async Task<PhoneCall> GetPhoneCallInformation(PhoneCallId id, CancellationToken token)
    {
        var request = new GetPhoneCallRequest(_service)
        {
            Id = id,
        };

        var response = await request.Execute(token).ConfigureAwait(false);
        return response.PhoneCall;
    }

    /// <summary>
    ///     Disconnects a phone call.
    /// </summary>
    /// <param name="id">The Id of the phone call.</param>
    /// <param name="token"></param>
    internal System.Threading.Tasks.Task DisconnectPhoneCall(PhoneCallId id, CancellationToken token)
    {
        var request = new DisconnectPhoneCallRequest(_service)
        {
            Id = id,
        };

        return request.Execute(token);
    }
}
