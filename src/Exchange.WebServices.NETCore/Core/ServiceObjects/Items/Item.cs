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
///     Represents a generic item. Properties available on items are defined in the ItemSchema class.
/// </summary>
[PublicAPI]
[Attachable]
[ServiceObjectDefinition(XmlElementNames.Item)]
public class Item : ServiceObject
{
    /// <summary>
    ///     Gets the parent attachment of this item.
    /// </summary>
    internal ItemAttachment ParentAttachment { get; }

    /// <summary>
    ///     Gets Id of the root item for this item.
    /// </summary>
    internal ItemId RootItemId
    {
        get
        {
            if (IsAttachment && ParentAttachment.Owner != null)
            {
                return ParentAttachment.Owner.RootItemId;
            }

            return Id;
        }
    }

    /// <summary>
    ///     Initializes an unsaved local instance of <see cref="Item" />. To bind to an existing item, use Item.Bind() instead.
    /// </summary>
    /// <param name="service">The ExchangeService object to which the item will be bound.</param>
    internal Item(ExchangeService service)
        : base(service)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="Item" /> class.
    /// </summary>
    /// <param name="parentAttachment">The parent attachment.</param>
    internal Item(ItemAttachment parentAttachment)
        : this(parentAttachment.Service)
    {
        EwsUtilities.Assert(parentAttachment != null, "Item.ctor", "parentAttachment is null");

        ParentAttachment = parentAttachment;
    }

    /// <summary>
    ///     Binds to an existing item, whatever its actual type is, and loads the specified set of properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to use to bind to the item.</param>
    /// <param name="id">The Id of the item to bind to.</param>
    /// <param name="propertySet">The set of properties to load.</param>
    /// <param name="token"></param>
    /// <returns>An Item instance representing the item corresponding to the specified Id.</returns>
    public static Task<Item> Bind(
        ExchangeService service,
        ItemId id,
        PropertySet propertySet,
        CancellationToken token = default
    )
    {
        return service.BindToItem<Item>(id, propertySet, token);
    }

    /// <summary>
    ///     Binds to an existing item, whatever its actual type is, and loads its first class properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to use to bind to the item.</param>
    /// <param name="id">The Id of the item to bind to.</param>
    /// <returns>An Item instance representing the item corresponding to the specified Id.</returns>
    public static Task<Item> Bind(ExchangeService service, ItemId id)
    {
        return Bind(service, id, PropertySet.FirstClassProperties);
    }

    /// <summary>
    ///     Internal method to return the schema associated with this type of object.
    /// </summary>
    /// <returns>The schema associated with this type of object.</returns>
    internal override ServiceObjectSchema GetSchema()
    {
        return ItemSchema.Instance;
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
    ///     Throws exception if this is attachment.
    /// </summary>
    internal void ThrowIfThisIsAttachment()
    {
        if (IsAttachment)
        {
            throw new InvalidOperationException(Strings.OperationDoesNotSupportAttachments);
        }
    }

    /// <summary>
    ///     The property definition for the Id of this object.
    /// </summary>
    /// <returns>A PropertyDefinition instance.</returns>
    internal override PropertyDefinition GetIdPropertyDefinition()
    {
        return ItemSchema.Id;
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
        ThrowIfThisIsNew();
        ThrowIfThisIsAttachment();

        return Service.InternalLoadPropertiesForItems([this,], propertySet, ServiceErrorHandling.ThrowOnError, token);
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
        return InternalDelete(deleteMode, sendCancellationsMode, affectedTaskOccurrences, false, token);
    }

    /// <summary>
    ///     Deletes the object.
    /// </summary>
    /// <param name="deleteMode">The deletion mode.</param>
    /// <param name="sendCancellationsMode">Indicates whether meeting cancellation messages should be sent.</param>
    /// <param name="affectedTaskOccurrences">Indicate which occurrence of a recurring task should be deleted.</param>
    /// <param name="suppressReadReceipts">Whether to suppress read receipts</param>
    /// <param name="token"></param>
    internal Task<ServiceResponseCollection<ServiceResponse>> InternalDelete(
        DeleteMode deleteMode,
        SendCancellationsMode? sendCancellationsMode,
        AffectedTaskOccurrence? affectedTaskOccurrences,
        bool suppressReadReceipts,
        CancellationToken token
    )
    {
        ThrowIfThisIsNew();
        ThrowIfThisIsAttachment();

        // If sendCancellationsMode is null, use the default value that's appropriate for item type.
        if (!sendCancellationsMode.HasValue)
        {
            sendCancellationsMode = DefaultSendCancellationsMode;
        }

        // If affectedTaskOccurrences is null, use the default value that's appropriate for item type.
        if (!affectedTaskOccurrences.HasValue)
        {
            affectedTaskOccurrences = DefaultAffectedTaskOccurrences;
        }

        return Service.DeleteItem(
            Id,
            deleteMode,
            sendCancellationsMode,
            affectedTaskOccurrences,
            suppressReadReceipts,
            token
        );
    }

    /// <summary>
    ///     Create item.
    /// </summary>
    /// <param name="parentFolderId">The parent folder id.</param>
    /// <param name="messageDisposition">The message disposition.</param>
    /// <param name="sendInvitationsMode">The send invitations mode.</param>
    /// <param name="token"></param>
    internal async System.Threading.Tasks.Task InternalCreate(
        FolderId? parentFolderId,
        MessageDisposition? messageDisposition,
        SendInvitationsMode? sendInvitationsMode,
        CancellationToken token
    )
    {
        ThrowIfThisIsNotNew();
        ThrowIfThisIsAttachment();

        if (IsNew || IsDirty)
        {
            await Service.CreateItem(
                    this,
                    parentFolderId,
                    messageDisposition,
                    sendInvitationsMode ?? DefaultSendInvitationsMode,
                    token
                )
                .ConfigureAwait(false);

            await Attachments.Save(token).ConfigureAwait(false);
        }
    }

    /// <summary>
    ///     Update item.
    /// </summary>
    /// <param name="parentFolderId">The parent folder id.</param>
    /// <param name="conflictResolutionMode">The conflict resolution mode.</param>
    /// <param name="messageDisposition">The message disposition.</param>
    /// <param name="sendInvitationsOrCancellationsMode">The send invitations or cancellations mode.</param>
    /// <param name="token"></param>
    /// <returns>Updated item.</returns>
    internal Task<Item?> InternalUpdate(
        FolderId? parentFolderId,
        ConflictResolutionMode conflictResolutionMode,
        MessageDisposition? messageDisposition,
        SendInvitationsOrCancellationsMode? sendInvitationsOrCancellationsMode,
        CancellationToken token
    )
    {
        return InternalUpdate(
            parentFolderId,
            conflictResolutionMode,
            messageDisposition,
            sendInvitationsOrCancellationsMode,
            suppressReadReceipts: false,
            token
        );
    }

    /// <summary>
    ///     Update item.
    /// </summary>
    /// <param name="parentFolderId">The parent folder id.</param>
    /// <param name="conflictResolutionMode">The conflict resolution mode.</param>
    /// <param name="messageDisposition">The message disposition.</param>
    /// <param name="sendInvitationsOrCancellationsMode">The send invitations or cancellations mode.</param>
    /// <param name="suppressReadReceipts">Whether to suppress read receipts</param>
    /// <param name="token"></param>
    /// <returns>Updated item.</returns>
    internal async Task<Item?> InternalUpdate(
        FolderId? parentFolderId,
        ConflictResolutionMode conflictResolutionMode,
        MessageDisposition? messageDisposition,
        SendInvitationsOrCancellationsMode? sendInvitationsOrCancellationsMode,
        bool suppressReadReceipts,
        CancellationToken token
    )
    {
        ThrowIfThisIsNew();
        ThrowIfThisIsAttachment();

        Item? returnedItem = null;

        if (IsDirty && PropertyBag.GetIsUpdateCallNecessary())
        {
            returnedItem = await Service.UpdateItem(
                    this,
                    parentFolderId,
                    conflictResolutionMode,
                    messageDisposition,
                    sendInvitationsOrCancellationsMode ?? DefaultSendInvitationsOrCancellationsMode,
                    suppressReadReceipts,
                    token
                )
                .ConfigureAwait(false);
        }

        // Regardless of whether item is dirty or not, if it has unprocessed
        // attachment changes, validate them and process now.
        if (HasUnprocessedAttachmentChanges())
        {
            Attachments.Validate();
            await Attachments.Save(token).ConfigureAwait(false);
        }

        return returnedItem;
    }

    /// <summary>
    ///     Gets a value indicating whether this instance has unprocessed attachment collection changes.
    /// </summary>
    internal bool HasUnprocessedAttachmentChanges()
    {
        return Attachments.HasUnprocessedChanges();
    }

    /// <summary>
    ///     Deletes the item. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="deleteMode">The deletion mode.</param>
    public Task<ServiceResponseCollection<ServiceResponse>> Delete(DeleteMode deleteMode)
    {
        return Delete(deleteMode, false);
    }

    /// <summary>
    ///     Deletes the item. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="deleteMode">The deletion mode.</param>
    /// <param name="suppressReadReceipts">Whether to suppress read receipts</param>
    /// <param name="token"></param>
    public Task<ServiceResponseCollection<ServiceResponse>> Delete(
        DeleteMode deleteMode,
        bool suppressReadReceipts,
        CancellationToken token = default
    )
    {
        return InternalDelete(deleteMode, null, null, suppressReadReceipts, token);
    }

    /// <summary>
    ///     Saves this item in a specific folder. Calling this method results in at least one call to EWS.
    ///     Multiple calls to EWS might be made if attachments have been added.
    /// </summary>
    /// <param name="parentFolderId">The Id of the folder in which to save this item.</param>
    /// <param name="token"></param>
    public System.Threading.Tasks.Task Save(FolderId parentFolderId, CancellationToken token = default)
    {
        EwsUtilities.ValidateParam(parentFolderId);

        return InternalCreate(parentFolderId, MessageDisposition.SaveOnly, null, token);
    }

    /// <summary>
    ///     Saves this item in a specific folder. Calling this method results in at least one call to EWS.
    ///     Multiple calls to EWS might be made if attachments have been added.
    /// </summary>
    /// <param name="parentFolderName">The name of the folder in which to save this item.</param>
    /// <param name="token"></param>
    public System.Threading.Tasks.Task Save(WellKnownFolderName parentFolderName, CancellationToken token = default)
    {
        return InternalCreate(new FolderId(parentFolderName), MessageDisposition.SaveOnly, null, token);
    }

    /// <summary>
    ///     Saves this item in the default folder based on the item's type (for example, an e-mail message is saved to the
    ///     Drafts folder).
    ///     Calling this method results in at least one call to EWS. Multiple calls to EWS might be made if attachments have
    ///     been added.
    /// </summary>
    public System.Threading.Tasks.Task Save(CancellationToken token = default)
    {
        return InternalCreate(null, MessageDisposition.SaveOnly, null, token);
    }

    /// <summary>
    ///     Applies the local changes that have been made to this item. Calling this method results in at least one call to
    ///     EWS.
    ///     Multiple calls to EWS might be made if attachments have been added or removed.
    /// </summary>
    /// <param name="conflictResolutionMode">The conflict resolution mode.</param>
    /// <param name="token"></param>
    public Task<Item?> Update(ConflictResolutionMode conflictResolutionMode, CancellationToken token = default)
    {
        return Update(conflictResolutionMode, false, token);
    }

    /// <summary>
    ///     Applies the local changes that have been made to this item. Calling this method results in at least one call to
    ///     EWS.
    ///     Multiple calls to EWS might be made if attachments have been added or removed.
    /// </summary>
    /// <param name="conflictResolutionMode">The conflict resolution mode.</param>
    /// <param name="suppressReadReceipts">Whether to suppress read receipts</param>
    /// <param name="token"></param>
    public Task<Item?> Update(
        ConflictResolutionMode conflictResolutionMode,
        bool suppressReadReceipts,
        CancellationToken token = default
    )
    {
        return InternalUpdate(
            null,
            conflictResolutionMode,
            MessageDisposition.SaveOnly,
            null,
            suppressReadReceipts,
            token
        );
    }

    /// <summary>
    ///     Creates a copy of this item in the specified folder. Calling this method results in a call to EWS.
    ///     <para>
    ///         Copy returns null if the copy operation is across two mailboxes or between a mailbox and a
    ///         public folder.
    ///     </para>
    /// </summary>
    /// <param name="destinationFolderId">The Id of the folder in which to create a copy of this item.</param>
    /// <param name="token"></param>
    /// <returns>The copy of this item.</returns>
    public Task<Item> Copy(FolderId destinationFolderId, CancellationToken token = default)
    {
        ThrowIfThisIsNew();
        ThrowIfThisIsAttachment();

        EwsUtilities.ValidateParam(destinationFolderId);

        return Service.CopyItem(Id, destinationFolderId, token);
    }

    /// <summary>
    ///     Creates a copy of this item in the specified folder. Calling this method results in a call to EWS.
    ///     <para>
    ///         Copy returns null if the copy operation is across two mailboxes or between a mailbox and a
    ///         public folder.
    ///     </para>
    /// </summary>
    /// <param name="destinationFolderName">The name of the folder in which to create a copy of this item.</param>
    /// <returns>The copy of this item.</returns>
    public Task<Item> Copy(WellKnownFolderName destinationFolderName)
    {
        return Copy(new FolderId(destinationFolderName));
    }

    /// <summary>
    ///     Moves this item to a the specified folder. Calling this method results in a call to EWS.
    ///     <para>
    ///         Move returns null if the move operation is across two mailboxes or between a mailbox and a
    ///         public folder.
    ///     </para>
    /// </summary>
    /// <param name="destinationFolderId">The Id of the folder to which to move this item.</param>
    /// <param name="token"></param>
    /// <returns>The moved copy of this item.</returns>
    public Task<Item> Move(FolderId destinationFolderId, CancellationToken token = default)
    {
        ThrowIfThisIsNew();
        ThrowIfThisIsAttachment();

        EwsUtilities.ValidateParam(destinationFolderId);

        return Service.MoveItem(Id, destinationFolderId, token);
    }

    /// <summary>
    ///     Moves this item to a the specified folder. Calling this method results in a call to EWS.
    ///     <para>
    ///         Move returns null if the move operation is across two mailboxes or between a mailbox and a
    ///         public folder.
    ///     </para>
    /// </summary>
    /// <param name="destinationFolderName">The name of the folder to which to move this item.</param>
    /// <returns>The moved copy of this item.</returns>
    public Task<Item> Move(WellKnownFolderName destinationFolderName)
    {
        return Move(new FolderId(destinationFolderName));
    }

    /// <summary>
    ///     Sets the extended property.
    /// </summary>
    /// <param name="extendedPropertyDefinition">The extended property definition.</param>
    /// <param name="value">The value.</param>
    public void SetExtendedProperty(ExtendedPropertyDefinition extendedPropertyDefinition, object value)
    {
        ExtendedProperties.SetExtendedProperty(extendedPropertyDefinition, value);
    }

    /// <summary>
    ///     Removes an extended property.
    /// </summary>
    /// <param name="extendedPropertyDefinition">The extended property definition.</param>
    /// <returns>True if property was removed.</returns>
    public bool RemoveExtendedProperty(ExtendedPropertyDefinition extendedPropertyDefinition)
    {
        return ExtendedProperties.RemoveExtendedProperty(extendedPropertyDefinition);
    }

    /// <summary>
    ///     Gets a list of extended properties defined on this object.
    /// </summary>
    /// <returns>Extended properties collection.</returns>
    internal override ExtendedPropertyCollection? GetExtendedProperties()
    {
        return ExtendedProperties;
    }

    /// <summary>
    ///     Validates this instance.
    /// </summary>
    internal override void Validate()
    {
        base.Validate();

        Attachments.Validate();

        // Flag parameter is only valid for Exchange2013 or higher
        if (TryGetProperty(ItemSchema.Flag, out Flag? flag) && flag != null)
        {
            if (Service.RequestedServerVersion < ExchangeVersion.Exchange2013)
            {
                throw new ServiceVersionException(
                    string.Format(Strings.ParameterIncompatibleWithRequestVersion, "Flag", ExchangeVersion.Exchange2013)
                );
            }

            flag.Validate();
        }
    }

    /// <summary>
    ///     Gets a value indicating whether a time zone SOAP header should be emitted in a CreateItem
    ///     or UpdateItem request so this item can be property saved or updated.
    /// </summary>
    /// <param name="isUpdateOperation">Indicates whether the operation being performed is an update operation.</param>
    /// <returns>
    ///     <c>true</c> if a time zone SOAP header should be emitted; otherwise, <c>false</c>.
    /// </returns>
    internal override bool GetIsTimeZoneHeaderRequired(bool isUpdateOperation)
    {
        // Starting E14SP2, attachment will be sent along with CreateItem requests. 
        // if the attachment used to require the Timezone header, CreateItem request should do so too.
        //
        if (!isUpdateOperation && Service.RequestedServerVersion >= ExchangeVersion.Exchange2010_SP2)
        {
            foreach (var itemAttachment in Attachments.OfType<ItemAttachment>())
            {
                if (itemAttachment.Item != null && itemAttachment.Item.GetIsTimeZoneHeaderRequired(false))
                {
                    return true;
                }
            }
        }

        return base.GetIsTimeZoneHeaderRequired(isUpdateOperation);
    }


    #region Properties

    /// <summary>
    ///     Gets a value indicating whether the item is an attachment.
    /// </summary>
    public bool IsAttachment => ParentAttachment != null;

    /// <summary>
    ///     Gets a value indicating whether this object is a real store item, or if it's a local object
    ///     that has yet to be saved.
    /// </summary>
    public override bool IsNew
    {
        get
        {
            // Item attachments don't have an Id, need to check whether the
            // parentAttachment is new or not.
            if (IsAttachment)
            {
                return ParentAttachment.IsNew;
            }

            return base.IsNew;
        }
    }

    /// <summary>
    ///     Gets the Id of this item.
    /// </summary>
    public ItemId Id => (ItemId)PropertyBag[GetIdPropertyDefinition()]!;

    /// <summary>
    ///     Get or sets the MIME content of this item.
    /// </summary>
    public MimeContent MimeContent
    {
        get => (MimeContent)PropertyBag[ItemSchema.MimeContent]!;
        set => PropertyBag[ItemSchema.MimeContent] = value;
    }

    /// <summary>
    ///     Get or sets the MimeContentUTF8 of this item.
    /// </summary>
    public MimeContentUTF8 MimeContentUTF8
    {
        get => (MimeContentUTF8)PropertyBag[ItemSchema.MimeContentUTF8]!;
        set => PropertyBag[ItemSchema.MimeContentUTF8] = value;
    }

    /// <summary>
    ///     Gets the Id of the parent folder of this item.
    /// </summary>
    public FolderId ParentFolderId => (FolderId)PropertyBag[ItemSchema.ParentFolderId]!;

    /// <summary>
    ///     Gets or sets the sensitivity of this item.
    /// </summary>
    public Sensitivity Sensitivity
    {
        get => (Sensitivity)PropertyBag[ItemSchema.Sensitivity]!;
        set => PropertyBag[ItemSchema.Sensitivity] = value;
    }

    /// <summary>
    ///     Gets a list of the attachments to this item.
    /// </summary>
    public AttachmentCollection Attachments => (AttachmentCollection)PropertyBag[ItemSchema.Attachments]!;

    /// <summary>
    ///     Gets the time when this item was received.
    /// </summary>
    public DateTime DateTimeReceived => (DateTime)PropertyBag[ItemSchema.DateTimeReceived]!;

    /// <summary>
    ///     Gets the size of this item.
    /// </summary>
    public int Size => (int)PropertyBag[ItemSchema.Size]!;

    /// <summary>
    ///     Gets or sets the list of categories associated with this item.
    /// </summary>
    public StringList Categories
    {
        get => (StringList)PropertyBag[ItemSchema.Categories]!;
        set => PropertyBag[ItemSchema.Categories] = value;
    }

    /// <summary>
    ///     Gets or sets the culture associated with this item.
    /// </summary>
    public string Culture
    {
        get => (string)PropertyBag[ItemSchema.Culture]!;
        set => PropertyBag[ItemSchema.Culture] = value;
    }

    /// <summary>
    ///     Gets or sets the importance of this item.
    /// </summary>
    public Importance Importance
    {
        get => (Importance)PropertyBag[ItemSchema.Importance]!;
        set => PropertyBag[ItemSchema.Importance] = value;
    }

    /// <summary>
    ///     Gets or sets the In-Reply-To reference of this item.
    /// </summary>
    public string InReplyTo
    {
        get => (string)PropertyBag[ItemSchema.InReplyTo]!;
        set => PropertyBag[ItemSchema.InReplyTo] = value;
    }

    /// <summary>
    ///     Gets a value indicating whether the message has been submitted to be sent.
    /// </summary>
    public bool IsSubmitted => (bool)PropertyBag[ItemSchema.IsSubmitted]!;

    /// <summary>
    ///     Gets a value indicating whether this is an associated item.
    /// </summary>
    public bool IsAssociated => (bool)PropertyBag[ItemSchema.IsAssociated]!;

    /// <summary>
    ///     Gets a value indicating whether the item is is a draft. An item is a draft when it has not yet been sent.
    /// </summary>
    public bool IsDraft => (bool)PropertyBag[ItemSchema.IsDraft]!;

    /// <summary>
    ///     Gets a value indicating whether the item has been sent by the current authenticated user.
    /// </summary>
    public bool IsFromMe => (bool)PropertyBag[ItemSchema.IsFromMe]!;

    /// <summary>
    ///     Gets a value indicating whether the item is a resend of another item.
    /// </summary>
    public bool IsResend => (bool)PropertyBag[ItemSchema.IsResend]!;

    /// <summary>
    ///     Gets a value indicating whether the item has been modified since it was created.
    /// </summary>
    public bool IsUnmodified => (bool)PropertyBag[ItemSchema.IsUnmodified]!;

    /// <summary>
    ///     Gets a list of Internet headers for this item.
    /// </summary>
    public InternetMessageHeaderCollection InternetMessageHeaders =>
        (InternetMessageHeaderCollection)PropertyBag[ItemSchema.InternetMessageHeaders]!;

    /// <summary>
    ///     Gets the date and time this item was sent.
    /// </summary>
    public DateTime DateTimeSent => (DateTime)PropertyBag[ItemSchema.DateTimeSent]!;

    /// <summary>
    ///     Gets the date and time this item was created.
    /// </summary>
    public DateTime DateTimeCreated => (DateTime)PropertyBag[ItemSchema.DateTimeCreated]!;

    /// <summary>
    ///     Gets a value indicating which response actions are allowed on this item. Examples of response actions are Reply and
    ///     Forward.
    /// </summary>
    public ResponseActions AllowedResponseActions => (ResponseActions)PropertyBag[ItemSchema.AllowedResponseActions]!;

    /// <summary>
    ///     Gets or sets the date and time when the reminder is due for this item.
    /// </summary>
    public DateTime ReminderDueBy
    {
        get => (DateTime)PropertyBag[ItemSchema.ReminderDueBy]!;
        set => PropertyBag[ItemSchema.ReminderDueBy] = value;
    }

    /// <summary>
    ///     Gets or sets a value indicating whether a reminder is set for this item.
    /// </summary>
    public bool IsReminderSet
    {
        get => (bool)PropertyBag[ItemSchema.IsReminderSet]!;
        set => PropertyBag[ItemSchema.IsReminderSet] = value;
    }

    /// <summary>
    ///     Gets or sets the number of minutes before the start of this item when the reminder should be triggered.
    /// </summary>
    public int ReminderMinutesBeforeStart
    {
        get => (int)PropertyBag[ItemSchema.ReminderMinutesBeforeStart]!;
        set => PropertyBag[ItemSchema.ReminderMinutesBeforeStart] = value;
    }

    /// <summary>
    ///     Gets a text summarizing the Cc recipients of this item.
    /// </summary>
    public string DisplayCc => (string)PropertyBag[ItemSchema.DisplayCc]!;

    /// <summary>
    ///     Gets a text summarizing the To recipients of this item.
    /// </summary>
    public string DisplayTo => (string)PropertyBag[ItemSchema.DisplayTo]!;

    /// <summary>
    ///     Gets a value indicating whether the item has attachments.
    /// </summary>
    public bool HasAttachments => (bool)PropertyBag[ItemSchema.HasAttachments]!;

    /// <summary>
    ///     Gets or sets the body of this item.
    /// </summary>
    public MessageBody Body
    {
        get => (MessageBody)PropertyBag[ItemSchema.Body]!;
        set => PropertyBag[ItemSchema.Body] = value;
    }

    /// <summary>
    ///     Gets or sets the custom class name of this item.
    /// </summary>
    public string ItemClass
    {
        get => (string)PropertyBag[ItemSchema.ItemClass]!;
        set => PropertyBag[ItemSchema.ItemClass] = value;
    }

    /// <summary>
    ///     Gets or sets the subject of this item.
    /// </summary>
    public string Subject
    {
        get => (string)PropertyBag[ItemSchema.Subject]!;
        set => SetSubject(value);
    }

    /// <summary>
    ///     Gets the query string that should be appended to the Exchange Web client URL to open this item using the
    ///     appropriate read form in a web browser.
    /// </summary>
    public string WebClientReadFormQueryString => (string)PropertyBag[ItemSchema.WebClientReadFormQueryString]!;

    /// <summary>
    ///     Gets the query string that should be appended to the Exchange Web client URL to open this item using the
    ///     appropriate edit form in a web browser.
    /// </summary>
    public string WebClientEditFormQueryString => (string)PropertyBag[ItemSchema.WebClientEditFormQueryString]!;

    /// <summary>
    ///     Gets a list of extended properties defined on this item.
    /// </summary>
    public ExtendedPropertyCollection ExtendedProperties =>
        (ExtendedPropertyCollection)PropertyBag[ServiceObjectSchema.ExtendedProperties]!;

    /// <summary>
    ///     Gets a value indicating the effective rights the current authenticated user has on this item.
    /// </summary>
    public EffectiveRights EffectiveRights => (EffectiveRights)PropertyBag[ItemSchema.EffectiveRights]!;

    /// <summary>
    ///     Gets the name of the user who last modified this item.
    /// </summary>
    public string LastModifiedName => (string)PropertyBag[ItemSchema.LastModifiedName]!;

    /// <summary>
    ///     Gets the date and time this item was last modified.
    /// </summary>
    public DateTime LastModifiedTime => (DateTime)PropertyBag[ItemSchema.LastModifiedTime]!;

    /// <summary>
    ///     Gets the Id of the conversation this item is part of.
    /// </summary>
    public ConversationId ConversationId => (ConversationId)PropertyBag[ItemSchema.ConversationId]!;

    /// <summary>
    ///     Gets the body part that is unique to the conversation this item is part of.
    /// </summary>
    public UniqueBody UniqueBody => (UniqueBody)PropertyBag[ItemSchema.UniqueBody]!;

    /// <summary>
    ///     Gets the store entry id.
    /// </summary>
    public byte[] StoreEntryId => (byte[])PropertyBag[ItemSchema.StoreEntryId]!;

    /// <summary>
    ///     Gets the item instance key.
    /// </summary>
    public byte[] InstanceKey => (byte[])PropertyBag[ItemSchema.InstanceKey]!;

    /// <summary>
    ///     Get or set the Flag value for this item.
    /// </summary>
    public Flag Flag
    {
        get => (Flag)PropertyBag[ItemSchema.Flag]!;
        set => PropertyBag[ItemSchema.Flag] = value;
    }

    /// <summary>
    ///     Gets the normalized body of the item.
    /// </summary>
    public NormalizedBody NormalizedBody => (NormalizedBody)PropertyBag[ItemSchema.NormalizedBody]!;

    /// <summary>
    ///     Gets the EntityExtractionResult of the item.
    /// </summary>
    public EntityExtractionResult EntityExtractionResult =>
        (EntityExtractionResult)PropertyBag[ItemSchema.EntityExtractionResult]!;

    /// <summary>
    ///     Gets or sets the policy tag.
    /// </summary>
    public PolicyTag PolicyTag
    {
        get => (PolicyTag)PropertyBag[ItemSchema.PolicyTag]!;
        set => PropertyBag[ItemSchema.PolicyTag] = value;
    }

    /// <summary>
    ///     Gets or sets the archive tag.
    /// </summary>
    public ArchiveTag ArchiveTag
    {
        get => (ArchiveTag)PropertyBag[ItemSchema.ArchiveTag]!;
        set => PropertyBag[ItemSchema.ArchiveTag] = value;
    }

    /// <summary>
    ///     Gets the retention date.
    /// </summary>
    public DateTime RetentionDate => (DateTime)PropertyBag[ItemSchema.RetentionDate]!;

    /// <summary>
    ///     Gets the item Preview.
    /// </summary>
    public string Preview => (string)PropertyBag[ItemSchema.Preview]!;

    /// <summary>
    ///     Gets the text body of the item.
    /// </summary>
    public TextBody TextBody => (TextBody)PropertyBag[ItemSchema.TextBody]!;

    /// <summary>
    ///     Gets the icon index.
    /// </summary>
    public IconIndex IconIndex => (IconIndex)PropertyBag[ItemSchema.IconIndex]!;

    /// <summary>
    ///     Gets or sets the list of hashtags associated with this item.
    /// </summary>
    public StringList Hashtags
    {
        get => (StringList)PropertyBag[ItemSchema.Hashtags]!;
        set => PropertyBag[ItemSchema.Hashtags] = value;
    }

    /// <summary>
    ///     Gets the Mentions associated with the message.
    /// </summary>
    public EmailAddressCollection Mentions
    {
        get => (EmailAddressCollection)PropertyBag[ItemSchema.Mentions]!;
        set => PropertyBag[ItemSchema.Mentions] = value;
    }

    /// <summary>
    ///     Gets a value indicating whether the item mentions me.
    /// </summary>
    public bool MentionedMe => (bool)PropertyBag[ItemSchema.MentionedMe]!;

    /// <summary>
    ///     Gets the default setting for how to treat affected task occurrences on Delete.
    ///     Subclasses will override this for different default behavior.
    /// </summary>
    internal virtual AffectedTaskOccurrence? DefaultAffectedTaskOccurrences => null;

    /// <summary>
    ///     Gets the default setting for sending cancellations on Delete.
    ///     Subclasses will override this for different default behavior.
    /// </summary>
    internal virtual SendCancellationsMode? DefaultSendCancellationsMode => null;

    /// <summary>
    ///     Gets the default settings for sending invitations on Save.
    ///     Subclasses will override this for different default behavior.
    /// </summary>
    internal virtual SendInvitationsMode? DefaultSendInvitationsMode => null;

    /// <summary>
    ///     Gets the default settings for sending invitations or cancellations on Update.
    ///     Subclasses will override this for different default behavior.
    /// </summary>
    internal virtual SendInvitationsOrCancellationsMode? DefaultSendInvitationsOrCancellationsMode => null;

    /// <summary>
    ///     Sets the subject.
    /// </summary>
    /// <param name="subject">The subject.</param>
    internal virtual void SetSubject(string subject)
    {
        PropertyBag[ItemSchema.Subject] = subject;
    }

    #endregion
}
