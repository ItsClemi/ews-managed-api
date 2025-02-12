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
///     Represents a reply to a post item.
/// </summary>
[PublicAPI]
[ServiceObjectDefinition(XmlElementNames.PostReplyItem, ReturnedByServer = false)]
public sealed class PostReply : ServiceObject
{
    private readonly Item _referenceItem;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PostReply" /> class.
    /// </summary>
    /// <param name="referenceItem">The reference item.</param>
    internal PostReply(Item referenceItem)
        : base(referenceItem.Service)
    {
        EwsUtilities.Assert(referenceItem != null, "PostReply.ctor", "referenceItem is null");

        referenceItem.ThrowIfThisIsNew();

        _referenceItem = referenceItem;
    }

    /// <summary>
    ///     Internal method to return the schema associated with this type of object.
    /// </summary>
    /// <returns>The schema associated with this type of object.</returns>
    internal override ServiceObjectSchema GetSchema()
    {
        return PostReplySchema.Instance;
    }

    /// <summary>
    ///     Gets the minimum required server version.
    /// </summary>
    /// <returns>Earliest Exchange version in which this service object type is supported.</returns>
    internal override ExchangeVersion GetMinimumRequiredServerVersion()
    {
        return ExchangeVersion.Exchange2007_SP1;
    }

    /// <summary>
    ///     Create a PostItem response.
    /// </summary>
    /// <param name="parentFolderId">The parent folder id.</param>
    /// <param name="messageDisposition">The message disposition.</param>
    /// <param name="token"></param>
    /// <returns>Created PostItem.</returns>
    internal async Task<PostItem> InternalCreate(
        FolderId? parentFolderId,
        MessageDisposition? messageDisposition,
        CancellationToken token
    )
    {
        ((ItemId)PropertyBag[ResponseObjectSchema.ReferenceItemId]).Assign(_referenceItem.Id);

        var items = await Service.InternalCreateResponseObject(this, parentFolderId, messageDisposition, token)
            .ConfigureAwait(false);

        var postItem = EwsUtilities.FindFirstItemOfType<PostItem>(items);

        // This should never happen. If it does, we have a bug.
        EwsUtilities.Assert(
            postItem != null,
            "PostReply.InternalCreate",
            "postItem is null. The CreateItem call did not return the expected PostItem."
        );

        return postItem;
    }

    /// <summary>
    ///     Loads the specified set of properties on the object.
    /// </summary>
    /// <param name="propertySet">The properties to load.</param>
    /// <param name="token"></param>
    internal override Task<ServiceResponseCollection<ServiceResponse>> InternalLoad(
        PropertySet propertySet,
        CancellationToken token
    )
    {
        throw new InvalidOperationException(Strings.LoadingThisObjectTypeNotSupported);
    }

    /// <summary>
    ///     Deletes the object.
    /// </summary>
    /// <param name="deleteMode">The deletion mode.</param>
    /// <param name="sendCancellationsMode">Indicates whether meeting cancellation messages should be sent.</param>
    /// <param name="affectedTaskOccurrences">Indicate which occurrence of a recurring task should be deleted.</param>
    /// <param name="token"></param>
    internal override Task<ServiceResponseCollection<ServiceResponse>> InternalDelete(
        DeleteMode deleteMode,
        SendCancellationsMode? sendCancellationsMode,
        AffectedTaskOccurrence? affectedTaskOccurrences,
        CancellationToken token
    )
    {
        throw new InvalidOperationException(Strings.DeletingThisObjectTypeNotAuthorized);
    }

    /// <summary>
    ///     Saves the post reply in the same folder as the original post item. Calling this method results in a call to EWS.
    /// </summary>
    /// <returns>A PostItem representing the posted reply.</returns>
    public async Task<PostItem> Save(CancellationToken token = default)
    {
        return await InternalCreate(null, null, token).ConfigureAwait(false);
    }

    /// <summary>
    ///     Saves the post reply in the specified folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="destinationFolderId">The Id of the folder in which to save the post reply.</param>
    /// <param name="token"></param>
    /// <returns>A PostItem representing the posted reply.</returns>
    public async Task<PostItem> Save(FolderId destinationFolderId, CancellationToken token = default)
    {
        EwsUtilities.ValidateParam(destinationFolderId);

        return await InternalCreate(destinationFolderId, null, token).ConfigureAwait(false);
    }

    /// <summary>
    ///     Saves the post reply in a specified folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="destinationFolderName">The name of the folder in which to save the post reply.</param>
    /// <param name="token"></param>
    /// <returns>A PostItem representing the posted reply.</returns>
    public async Task<PostItem> Save(WellKnownFolderName destinationFolderName, CancellationToken token = default)
    {
        return await InternalCreate(new FolderId(destinationFolderName), null, token).ConfigureAwait(false);
    }


    #region Properties

    /// <summary>
    ///     Gets or sets the subject of the post reply.
    /// </summary>
    public string Subject
    {
        get => (string)PropertyBag[ItemSchema.Subject];
        set => PropertyBag[ItemSchema.Subject] = value;
    }

    /// <summary>
    ///     Gets or sets the body of the post reply.
    /// </summary>
    public MessageBody Body
    {
        get => (MessageBody)PropertyBag[ItemSchema.Body];
        set => PropertyBag[ItemSchema.Body] = value;
    }

    /// <summary>
    ///     Gets or sets the body prefix that should be prepended to the original post item's body.
    /// </summary>
    public MessageBody BodyPrefix
    {
        get => (MessageBody)PropertyBag[ResponseObjectSchema.BodyPrefix];
        set => PropertyBag[ResponseObjectSchema.BodyPrefix] = value;
    }

    #endregion
}
