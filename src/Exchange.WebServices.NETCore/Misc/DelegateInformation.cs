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

using System.Collections.ObjectModel;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents the results of a GetDelegates operation.
/// </summary>
[PublicAPI]
public sealed class DelegateInformation
{
    /// <summary>
    ///     Gets a list of responses for each of the delegate users concerned by the operation.
    /// </summary>
    public Collection<DelegateUserResponse> DelegateUserResponses { get; }

    /// <summary>
    ///     Gets a value indicating if and how meeting requests are delivered to delegates.
    /// </summary>
    public MeetingRequestsDeliveryScope MeetingRequestsDeliveryScope { get; }

    /// <summary>
    ///     Initializes a DelegateInformation object
    /// </summary>
    /// <param name="delegateUserResponses">List of DelegateUserResponses from a GetDelegates request</param>
    /// <param name="meetingRequestsDeliveryScope">MeetingRequestsDeliveryScope from a GetDelegates request.</param>
    internal DelegateInformation(
        IList<DelegateUserResponse> delegateUserResponses,
        MeetingRequestsDeliveryScope meetingRequestsDeliveryScope
    )
    {
        DelegateUserResponses = new Collection<DelegateUserResponse>(delegateUserResponses);
        MeetingRequestsDeliveryScope = meetingRequestsDeliveryScope;
    }
}
