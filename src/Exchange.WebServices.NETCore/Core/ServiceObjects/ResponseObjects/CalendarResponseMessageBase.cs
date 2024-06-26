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

using System.ComponentModel;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents the base class for all calendar-related response messages.
/// </summary>
/// <typeparam name="TMessage">The type of message that is created when this response message is saved.</typeparam>
[PublicAPI]
[EditorBrowsable(EditorBrowsableState.Never)]
public abstract class CalendarResponseMessageBase<TMessage> : ResponseObject<TMessage>
    where TMessage : EmailMessage
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="CalendarResponseMessageBase&lt;TMessage&gt;" /> class.
    /// </summary>
    /// <param name="referenceItem">The reference item.</param>
    internal CalendarResponseMessageBase(Item referenceItem)
        : base(referenceItem)
    {
    }

    /// <summary>
    ///     Saves the response in the specified folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="destinationFolderId">The Id of the folder in which to save the response.</param>
    /// <param name="token"></param>
    /// <returns>
    ///     A CalendarActionResults object containing the various items that were created or modified as a
    ///     results of this operation.
    /// </returns>
    public new async Task<CalendarActionResults> Save(FolderId destinationFolderId, CancellationToken token = default)
    {
        EwsUtilities.ValidateParam(destinationFolderId);

        return new CalendarActionResults(
            await InternalCreate(destinationFolderId, MessageDisposition.SaveOnly, token).ConfigureAwait(false)
        );
    }

    /// <summary>
    ///     Saves the response in the specified folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="destinationFolderName">The name of the folder in which to save the response.</param>
    /// <param name="token"></param>
    /// <returns>
    ///     A CalendarActionResults object containing the various items that were created or modified as a
    ///     results of this operation.
    /// </returns>
    public new async Task<CalendarActionResults> Save(
        WellKnownFolderName destinationFolderName,
        CancellationToken token = default
    )
    {
        var result = await InternalCreate(new FolderId(destinationFolderName), MessageDisposition.SaveOnly, token)
            .ConfigureAwait(false);

        return new CalendarActionResults(result);
    }

    /// <summary>
    ///     Saves the response in the Drafts folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <returns>
    ///     A CalendarActionResults object containing the various items that were created or modified as a
    ///     results of this operation.
    /// </returns>
    public new async Task<CalendarActionResults> Save(CancellationToken token = default)
    {
        var result = await InternalCreate(null, MessageDisposition.SaveOnly, token).ConfigureAwait(false);
        return new CalendarActionResults(result);
    }

    /// <summary>
    ///     Sends this response without saving a copy. Calling this method results in a call to EWS.
    /// </summary>
    /// <returns>
    ///     A CalendarActionResults object containing the various items that were created or modified as a
    ///     results of this operation.
    /// </returns>
    public new async Task<CalendarActionResults> Send(CancellationToken token = default)
    {
        var result = await InternalCreate(null, MessageDisposition.SendOnly, token).ConfigureAwait(false);
        return new CalendarActionResults(result);
    }

    /// <summary>
    ///     Sends this response ans saves a copy in the specified folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="destinationFolderId">The Id of the folder in which to save the copy of the message.</param>
    /// <param name="token"></param>
    /// <returns>
    ///     A CalendarActionResults object containing the various items that were created or modified as a
    ///     results of this operation.
    /// </returns>
    public new async Task<CalendarActionResults> SendAndSaveCopy(
        FolderId destinationFolderId,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(destinationFolderId);

        return new CalendarActionResults(
            await InternalCreate(destinationFolderId, MessageDisposition.SendAndSaveCopy, token).ConfigureAwait(false)
        );
    }

    /// <summary>
    ///     Sends this response and saves a copy in the specified folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="destinationFolderName">The name of the folder in which to save the copy of the message.</param>
    /// <param name="token"></param>
    /// <returns>
    ///     A CalendarActionResults object containing the various items that were created or modified as a
    ///     results of this operation.
    /// </returns>
    public new async Task<CalendarActionResults> SendAndSaveCopy(
        WellKnownFolderName destinationFolderName,
        CancellationToken token = default
    )
    {
        var result = await InternalCreate(
                new FolderId(destinationFolderName),
                MessageDisposition.SendAndSaveCopy,
                token
            )
            .ConfigureAwait(false);

        return new CalendarActionResults(result);
    }

    /// <summary>
    ///     Sends this response ans saves a copy in the Sent Items folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <returns>
    ///     A CalendarActionResults object containing the various items that were created or modified as a
    ///     results of this operation.
    /// </returns>
    public new async Task<CalendarActionResults> SendAndSaveCopy(CancellationToken token = default)
    {
        var result = await InternalCreate(null, MessageDisposition.SendAndSaveCopy, token).ConfigureAwait(false);

        return new CalendarActionResults(result);
    }
}
