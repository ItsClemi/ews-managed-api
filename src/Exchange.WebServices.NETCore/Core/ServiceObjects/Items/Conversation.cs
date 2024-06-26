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
///     Represents a collection of Conversation related properties.
///     Properties available on this object are defined in the ConversationSchema class.
/// </summary>
[PublicAPI]
[ServiceObjectDefinition(XmlElementNames.Conversation)]
public class Conversation : ServiceObject
{
    /// <summary>
    ///     Initializes an unsaved local instance of <see cref="Conversation" />.
    /// </summary>
    /// <param name="service">The ExchangeService object to which the item will be bound.</param>
    internal Conversation(ExchangeService service)
        : base(service)
    {
    }

    /// <summary>
    ///     Internal method to return the schema associated with this type of object.
    /// </summary>
    /// <returns>The schema associated with this type of object.</returns>
    internal override ServiceObjectSchema GetSchema()
    {
        return ConversationSchema.Instance;
    }

    /// <summary>
    ///     Gets the minimum required server version.
    /// </summary>
    /// <returns>Earliest Exchange version in which this service object type is supported.</returns>
    internal override ExchangeVersion GetMinimumRequiredServerVersion()
    {
        return ExchangeVersion.Exchange2010_SP1;
    }

    /// <summary>
    ///     The property definition for the Id of this object.
    /// </summary>
    /// <returns>A PropertyDefinition instance.</returns>
    internal override PropertyDefinition GetIdPropertyDefinition()
    {
        return ConversationSchema.Id;
    }


    #region Not Supported Methods or properties

    /// <summary>
    ///     This method is not supported in this object.
    ///     Loads the specified set of properties on the object.
    /// </summary>
    /// <param name="propertySet">The properties to load.</param>
    /// <param name="token"></param>
    internal override Task<ServiceResponseCollection<ServiceResponse>> InternalLoad(
        PropertySet propertySet,
        CancellationToken token
    )
    {
        throw new NotSupportedException();
    }

    /// <summary>
    ///     This is not supported in this object.
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
        throw new NotSupportedException();
    }

    /// <summary>
    ///     This method is not supported in this object.
    ///     Gets the name of the change XML element.
    /// </summary>
    /// <returns>XML element name,</returns>
    internal override string GetChangeXmlElementName()
    {
        throw new NotSupportedException();
    }

    /// <summary>
    ///     This method is not supported in this object.
    ///     Gets the name of the delete field XML element.
    /// </summary>
    /// <returns>XML element name,</returns>
    internal override string GetDeleteFieldXmlElementName()
    {
        throw new NotSupportedException();
    }

    /// <summary>
    ///     This method is not supported in this object.
    ///     Gets the name of the set field XML element.
    /// </summary>
    /// <returns>XML element name,</returns>
    internal override string GetSetFieldXmlElementName()
    {
        throw new NotSupportedException();
    }

    /// <summary>
    ///     This method is not supported in this object.
    ///     Gets a value indicating whether a time zone SOAP header should be emitted in a CreateItem
    ///     or UpdateItem request so this item can be property saved or updated.
    /// </summary>
    /// <param name="isUpdateOperation">Indicates whether the operation being petrformed is an update operation.</param>
    /// <returns><c>true</c> if a time zone SOAP header should be emitted; otherwise, <c>false</c>.</returns>
    internal override bool GetIsTimeZoneHeaderRequired(bool isUpdateOperation)
    {
        throw new NotSupportedException();
    }

    #endregion


    #region Conversation Action Methods

    /// <summary>
    ///     Sets up a conversation so that any item received within that conversation is always categorized.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="categories">The categories that should be stamped on items in the conversation.</param>
    /// <param name="processSynchronously">
    ///     Indicates whether the method should return only once enabling this rule and stamping existing items
    ///     in the conversation is completely done. If processSynchronously is false, the method returns immediately.
    /// </param>
    public async System.Threading.Tasks.Task EnableAlwaysCategorizeItems(
        IEnumerable<string> categories,
        bool processSynchronously
    )
    {
        var responses = await Service.EnableAlwaysCategorizeItemsInConversations(
                new[]
                {
                    Id,
                },
                categories,
                processSynchronously
            )
            .ConfigureAwait(false);

        responses[0].ThrowIfNecessary();
    }

    /// <summary>
    ///     Sets up a conversation so that any item received within that conversation is no longer categorized.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="processSynchronously">
    ///     Indicates whether the method should return only once disabling this rule and removing the categories from existing
    ///     items
    ///     in the conversation is completely done. If processSynchronously is false, the method returns immediately.
    /// </param>
    public async System.Threading.Tasks.Task DisableAlwaysCategorizeItems(bool processSynchronously)
    {
        var responses = await Service.DisableAlwaysCategorizeItemsInConversations(
                new[]
                {
                    Id,
                },
                processSynchronously
            )
            .ConfigureAwait(false);

        responses[0].ThrowIfNecessary();
    }

    /// <summary>
    ///     Sets up a conversation so that any item received within that conversation is always moved to Deleted Items folder.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="processSynchronously">
    ///     Indicates whether the method should return only once enabling this rule and deleting existing items
    ///     in the conversation is completely done. If processSynchronously is false, the method returns immediately.
    /// </param>
    public async System.Threading.Tasks.Task EnableAlwaysDeleteItems(bool processSynchronously)
    {
        var responses = await Service.EnableAlwaysDeleteItemsInConversations(
                new[]
                {
                    Id,
                },
                processSynchronously
            )
            .ConfigureAwait(false);

        responses[0].ThrowIfNecessary();
    }

    /// <summary>
    ///     Sets up a conversation so that any item received within that conversation is no longer moved to Deleted Items
    ///     folder.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="processSynchronously">
    ///     Indicates whether the method should return only once disabling this rule and restoring the items
    ///     in the conversation is completely done. If processSynchronously is false, the method returns immediately.
    /// </param>
    public async System.Threading.Tasks.Task DisableAlwaysDeleteItems(bool processSynchronously)
    {
        var responses = await Service.DisableAlwaysDeleteItemsInConversations(
                new[]
                {
                    Id,
                },
                processSynchronously
            )
            .ConfigureAwait(false);
        responses[0].ThrowIfNecessary();
    }

    /// <summary>
    ///     Sets up a conversation so that any item received within that conversation is always moved to a specific folder.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="destinationFolderId">The Id of the folder to which conversation items should be moved.</param>
    /// <param name="processSynchronously">
    ///     Indicates whether the method should return only once enabling this rule
    ///     and moving existing items in the conversation is completely done. If processSynchronously is false, the method
    ///     returns immediately.
    /// </param>
    public async System.Threading.Tasks.Task EnableAlwaysMoveItems(
        FolderId destinationFolderId,
        bool processSynchronously
    )
    {
        var responses = await Service.EnableAlwaysMoveItemsInConversations(
                new[]
                {
                    Id,
                },
                destinationFolderId,
                processSynchronously
            )
            .ConfigureAwait(false);

        responses[0].ThrowIfNecessary();
    }

    /// <summary>
    ///     Sets up a conversation so that any item received within that conversation is no longer moved to a specific
    ///     folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="processSynchronously">
    ///     Indicates whether the method should return only once disabling this
    ///     rule is completely done. If processSynchronously is false, the method returns immediately.
    /// </param>
    public async System.Threading.Tasks.Task DisableAlwaysMoveItemsInConversation(bool processSynchronously)
    {
        var responses = await Service.DisableAlwaysMoveItemsInConversations(
                new[]
                {
                    Id,
                },
                processSynchronously
            )
            .ConfigureAwait(false);

        responses[0].ThrowIfNecessary();
    }

    /// <summary>
    ///     Deletes items in the specified conversation.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="contextFolderId">
    ///     The Id of the folder items must belong to in order to be deleted. If contextFolderId is
    ///     null, items across the entire mailbox are deleted.
    /// </param>
    /// <param name="deleteMode">The deletion mode.</param>
    public async System.Threading.Tasks.Task DeleteItems(FolderId contextFolderId, DeleteMode deleteMode)
    {
        var responses = await Service.DeleteItemsInConversations(
                new[]
                {
                    new KeyValuePair<ConversationId, DateTime?>(Id, GlobalLastDeliveryTime),
                },
                contextFolderId,
                deleteMode
            )
            .ConfigureAwait(false);

        responses[0].ThrowIfNecessary();
    }

    /// <summary>
    ///     Moves items in the specified conversation to a specific folder.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="contextFolderId">
    ///     The Id of the folder items must belong to in order to be moved. If contextFolderId is null,
    ///     items across the entire mailbox are moved.
    /// </param>
    /// <param name="destinationFolderId">The Id of the destination folder.</param>
    public async System.Threading.Tasks.Task MoveItemsInConversation(
        FolderId contextFolderId,
        FolderId destinationFolderId
    )
    {
        var responses = await Service.MoveItemsInConversations(
                new[]
                {
                    new KeyValuePair<ConversationId, DateTime?>(Id, GlobalLastDeliveryTime),
                },
                contextFolderId,
                destinationFolderId
            )
            .ConfigureAwait(false);

        responses[0].ThrowIfNecessary();
    }

    /// <summary>
    ///     Copies items in the specified conversation to a specific folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="contextFolderId">
    ///     The Id of the folder items must belong to in order to be copied. If contextFolderId
    ///     is null, items across the entire mailbox are copied.
    /// </param>
    /// <param name="destinationFolderId">The Id of the destination folder.</param>
    public async System.Threading.Tasks.Task CopyItemsInConversation(
        FolderId contextFolderId,
        FolderId destinationFolderId
    )
    {
        var responses = await Service.CopyItemsInConversations(
                new[]
                {
                    new KeyValuePair<ConversationId, DateTime?>(Id, GlobalLastDeliveryTime),
                },
                contextFolderId,
                destinationFolderId
            )
            .ConfigureAwait(false);

        responses[0].ThrowIfNecessary();
    }

    /// <summary>
    ///     Sets the read state of items in the specified conversation. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="contextFolderId">
    ///     The Id of the folder items must belong to in order for their read state to
    ///     be set. If contextFolderId is null, the read states of items across the entire mailbox are set.
    /// </param>
    /// <param name="isRead">
    ///     if set to <c>true</c>, conversation items are marked as read; otherwise they are
    ///     marked as unread.
    /// </param>
    public async System.Threading.Tasks.Task SetReadStateForItemsInConversation(FolderId contextFolderId, bool isRead)
    {
        var responses = await Service.SetReadStateForItemsInConversations(
                new[]
                {
                    new KeyValuePair<ConversationId, DateTime?>(Id, GlobalLastDeliveryTime),
                },
                contextFolderId,
                isRead
            )
            .ConfigureAwait(false);

        responses[0].ThrowIfNecessary();
    }

    /// <summary>
    ///     Sets the read state of items in the specified conversation. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="contextFolderId">
    ///     The Id of the folder items must belong to in order for their read state to
    ///     be set. If contextFolderId is null, the read states of items across the entire mailbox are set.
    /// </param>
    /// <param name="isRead">
    ///     if set to <c>true</c>, conversation items are marked as read; otherwise they are
    ///     marked as unread.
    /// </param>
    /// <param name="suppressReadReceipts">if set to <c>true</c> read receipts are suppressed.</param>
    public async System.Threading.Tasks.Task SetReadStateForItemsInConversation(
        FolderId contextFolderId,
        bool isRead,
        bool suppressReadReceipts
    )
    {
        var responses = await Service.SetReadStateForItemsInConversations(
                new[]
                {
                    new KeyValuePair<ConversationId, DateTime?>(Id, GlobalLastDeliveryTime),
                },
                contextFolderId,
                isRead,
                suppressReadReceipts
            )
            .ConfigureAwait(false);

        responses[0].ThrowIfNecessary();
    }

    /// <summary>
    ///     Sets the retention policy of items in the specified conversation. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="contextFolderId">
    ///     The Id of the folder items must belong to in order for their retention policy to
    ///     be set. If contextFolderId is null, the retention policy of items across the entire mailbox are set.
    /// </param>
    /// <param name="retentionPolicyType">Retention policy type.</param>
    /// <param name="retentionPolicyTagId">Retention policy tag id.  Null will clear the policy.</param>
    public async System.Threading.Tasks.Task SetRetentionPolicyForItemsInConversation(
        FolderId contextFolderId,
        RetentionType retentionPolicyType,
        Guid? retentionPolicyTagId
    )
    {
        var responses = await Service.SetRetentionPolicyForItemsInConversations(
                new[]
                {
                    new KeyValuePair<ConversationId, DateTime?>(Id, GlobalLastDeliveryTime),
                },
                contextFolderId,
                retentionPolicyType,
                retentionPolicyTagId
            )
            .ConfigureAwait(false);

        responses[0].ThrowIfNecessary();
    }

    /// <summary>
    ///     Flag conversation items as complete. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="contextFolderId">
    ///     The Id of the folder items must belong to in order to be flagged as complete. If contextFolderId is
    ///     null, items in conversation across the entire mailbox are marked as complete.
    /// </param>
    /// <param name="completeDate">The complete date (can be null).</param>
    public async System.Threading.Tasks.Task FlagItemsComplete(FolderId contextFolderId, DateTime? completeDate)
    {
        var flag = new Flag
        {
            FlagStatus = ItemFlagStatus.Complete,
        };

        if (completeDate.HasValue)
        {
            flag.CompleteDate = completeDate.Value;
        }

        var responses = await Service.SetFlagStatusForItemsInConversations(
                new[]
                {
                    new KeyValuePair<ConversationId, DateTime?>(Id, GlobalLastDeliveryTime),
                },
                contextFolderId,
                flag
            )
            .ConfigureAwait(false);

        responses[0].ThrowIfNecessary();
    }

    /// <summary>
    ///     Clear flags for conversation items. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="contextFolderId">
    ///     The Id of the folder items must belong to in order to be unflagged. If contextFolderId is
    ///     null, flags for items in conversation across the entire mailbox are cleared.
    /// </param>
    public async System.Threading.Tasks.Task ClearItemFlags(FolderId contextFolderId)
    {
        var flag = new Flag
        {
            FlagStatus = ItemFlagStatus.NotFlagged,
        };

        var responses = await Service.SetFlagStatusForItemsInConversations(
                new[]
                {
                    new KeyValuePair<ConversationId, DateTime?>(Id, GlobalLastDeliveryTime),
                },
                contextFolderId,
                flag
            )
            .ConfigureAwait(false);

        responses[0].ThrowIfNecessary();
    }

    /// <summary>
    ///     Flags conversation items. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="contextFolderId">
    ///     The Id of the folder items must belong to in order to be flagged. If contextFolderId is
    ///     null, items in conversation across the entire mailbox are flagged.
    /// </param>
    /// <param name="startDate">The start date (can be null).</param>
    /// <param name="dueDate">The due date (can be null).</param>
    public async System.Threading.Tasks.Task FlagItems(FolderId contextFolderId, DateTime? startDate, DateTime? dueDate)
    {
        var flag = new Flag
        {
            FlagStatus = ItemFlagStatus.Flagged,
        };
        if (startDate.HasValue)
        {
            flag.StartDate = startDate.Value;
        }

        if (dueDate.HasValue)
        {
            flag.DueDate = dueDate.Value;
        }

        var responses = await Service.SetFlagStatusForItemsInConversations(
                new[]
                {
                    new KeyValuePair<ConversationId, DateTime?>(Id, GlobalLastDeliveryTime),
                },
                contextFolderId,
                flag
            )
            .ConfigureAwait(false);

        responses[0].ThrowIfNecessary();
    }

    #endregion


    #region Properties

    /// <summary>
    ///     Gets the Id of this Conversation.
    /// </summary>
    public ConversationId Id => (ConversationId)PropertyBag[GetIdPropertyDefinition()];

    /// <summary>
    ///     Gets the topic of this Conversation.
    /// </summary>
    public string Topic
    {
        get
        {
            var returnValue = string.Empty;

            // This property need not be present hence the property bag may not contain it.
            // Check for the presence of this property before accessing it.
            if (PropertyBag.Contains(ConversationSchema.Topic))
            {
                PropertyBag.TryGetProperty(ConversationSchema.Topic, out returnValue);
            }

            return returnValue;
        }
    }

    /// <summary>
    ///     Gets a list of all the people who have received messages in this conversation in the current folder only.
    /// </summary>
    public StringList UniqueRecipients => (StringList)PropertyBag[ConversationSchema.UniqueRecipients];

    /// <summary>
    ///     Gets a list of all the people who have received messages in this conversation across all folders in the mailbox.
    /// </summary>
    public StringList GlobalUniqueRecipients => (StringList)PropertyBag[ConversationSchema.GlobalUniqueRecipients];

    /// <summary>
    ///     Gets a list of all the people who have sent messages that are currently unread in this conversation in the current
    ///     folder only.
    /// </summary>
    public StringList UniqueUnreadSenders
    {
        get
        {
            StringList? unreadSenders = null;

            // This property need not be present hence the property bag may not contain it.
            // Check for the presence of this property before accessing it.
            if (PropertyBag.Contains(ConversationSchema.UniqueUnreadSenders))
            {
                PropertyBag.TryGetProperty(ConversationSchema.UniqueUnreadSenders, out unreadSenders);
            }

            return unreadSenders;
        }
    }

    /// <summary>
    ///     Gets a list of all the people who have sent messages that are currently unread in this conversation across all
    ///     folders in the mailbox.
    /// </summary>
    public StringList GlobalUniqueUnreadSenders
    {
        get
        {
            StringList? unreadSenders = null;

            // This property need not be present hence the property bag may not contain it.
            // Check for the presence of this property before accessing it.
            if (PropertyBag.Contains(ConversationSchema.GlobalUniqueUnreadSenders))
            {
                PropertyBag.TryGetProperty(ConversationSchema.GlobalUniqueUnreadSenders, out unreadSenders);
            }

            return unreadSenders;
        }
    }

    /// <summary>
    ///     Gets a list of all the people who have sent messages in this conversation in the current folder only.
    /// </summary>
    public StringList UniqueSenders => (StringList)PropertyBag[ConversationSchema.UniqueSenders];

    /// <summary>
    ///     Gets a list of all the people who have sent messages in this conversation across all folders in the mailbox.
    /// </summary>
    public StringList GlobalUniqueSenders => (StringList)PropertyBag[ConversationSchema.GlobalUniqueSenders];

    /// <summary>
    ///     Gets the delivery time of the message that was last received in this conversation in the current folder only.
    /// </summary>
    public DateTime LastDeliveryTime => (DateTime)PropertyBag[ConversationSchema.LastDeliveryTime];

    /// <summary>
    ///     Gets the delivery time of the message that was last received in this conversation across all folders in the
    ///     mailbox.
    /// </summary>
    public DateTime GlobalLastDeliveryTime => (DateTime)PropertyBag[ConversationSchema.GlobalLastDeliveryTime];

    /// <summary>
    ///     Gets a list summarizing the categories stamped on messages in this conversation, in the current folder only.
    /// </summary>
    public StringList Categories
    {
        get
        {
            StringList? returnValue = null;

            // This property need not be present hence the property bag may not contain it.
            // Check for the presence of this property before accessing it.
            if (PropertyBag.Contains(ConversationSchema.Categories))
            {
                PropertyBag.TryGetProperty(ConversationSchema.Categories, out returnValue);
            }

            return returnValue;
        }
    }

    /// <summary>
    ///     Gets a list summarizing the categories stamped on messages in this conversation, across all folders in the mailbox.
    /// </summary>
    public StringList GlobalCategories
    {
        get
        {
            StringList? returnValue = null;

            // This property need not be present hence the property bag may not contain it.
            // Check for the presence of this property before accessing it.
            if (PropertyBag.Contains(ConversationSchema.GlobalCategories))
            {
                PropertyBag.TryGetProperty(ConversationSchema.GlobalCategories, out returnValue);
            }

            return returnValue;
        }
    }

    /// <summary>
    ///     Gets the flag status for this conversation, calculated by aggregating individual messages flag status in the
    ///     current folder.
    /// </summary>
    public ConversationFlagStatus FlagStatus
    {
        get
        {
            var returnValue = ConversationFlagStatus.NotFlagged;

            // This property need not be present hence the property bag may not contain it.
            // Check for the presence of this property before accessing it.
            if (PropertyBag.Contains(ConversationSchema.FlagStatus))
            {
                PropertyBag.TryGetProperty(ConversationSchema.FlagStatus, out returnValue);
            }

            return returnValue;
        }
    }

    /// <summary>
    ///     Gets the flag status for this conversation, calculated by aggregating individual messages flag status across all
    ///     folders in the mailbox.
    /// </summary>
    public ConversationFlagStatus GlobalFlagStatus
    {
        get
        {
            var returnValue = ConversationFlagStatus.NotFlagged;

            // This property need not be present hence the property bag may not contain it.
            // Check for the presence of this property before accessing it.
            if (PropertyBag.Contains(ConversationSchema.GlobalFlagStatus))
            {
                PropertyBag.TryGetProperty(ConversationSchema.GlobalFlagStatus, out returnValue);
            }

            return returnValue;
        }
    }

    /// <summary>
    ///     Gets a value indicating if at least one message in this conversation, in the current folder only, has an
    ///     attachment.
    /// </summary>
    public bool HasAttachments => (bool)PropertyBag[ConversationSchema.HasAttachments];

    /// <summary>
    ///     Gets a value indicating if at least one message in this conversation, across all folders in the mailbox, has an
    ///     attachment.
    /// </summary>
    public bool GlobalHasAttachments => (bool)PropertyBag[ConversationSchema.GlobalHasAttachments];

    /// <summary>
    ///     Gets the total number of messages in this conversation in the current folder only.
    /// </summary>
    public int MessageCount => (int)PropertyBag[ConversationSchema.MessageCount];

    /// <summary>
    ///     Gets the total number of messages in this conversation across all folders in the mailbox.
    /// </summary>
    public int GlobalMessageCount => (int)PropertyBag[ConversationSchema.GlobalMessageCount];

    /// <summary>
    ///     Gets the total number of unread messages in this conversation in the current folder only.
    /// </summary>
    public int UnreadCount
    {
        get
        {
            var returnValue = 0;

            // This property need not be present hence the property bag may not contain it.
            // Check for the presence of this property before accessing it.
            if (PropertyBag.Contains(ConversationSchema.UnreadCount))
            {
                PropertyBag.TryGetProperty(ConversationSchema.UnreadCount, out returnValue);
            }

            return returnValue;
        }
    }

    /// <summary>
    ///     Gets the total number of unread messages in this conversation across all folders in the mailbox.
    /// </summary>
    public int GlobalUnreadCount
    {
        get
        {
            var returnValue = 0;

            // This property need not be present hence the property bag may not contain it.
            // Check for the presence of this property before accessing it.
            if (PropertyBag.Contains(ConversationSchema.GlobalUnreadCount))
            {
                PropertyBag.TryGetProperty(ConversationSchema.GlobalUnreadCount, out returnValue);
            }

            return returnValue;
        }
    }

    /// <summary>
    ///     Gets the size of this conversation, calculated by adding the sizes of all messages in the conversation in the
    ///     current folder only.
    /// </summary>
    public int Size => (int)PropertyBag[ConversationSchema.Size];

    /// <summary>
    ///     Gets the size of this conversation, calculated by adding the sizes of all messages in the conversation across all
    ///     folders in the mailbox.
    /// </summary>
    public int GlobalSize => (int)PropertyBag[ConversationSchema.GlobalSize];

    /// <summary>
    ///     Gets a list summarizing the classes of the items in this conversation, in the current folder only.
    /// </summary>
    public StringList ItemClasses => (StringList)PropertyBag[ConversationSchema.ItemClasses];

    /// <summary>
    ///     Gets a list summarizing the classes of the items in this conversation, across all folders in the mailbox.
    /// </summary>
    public StringList GlobalItemClasses => (StringList)PropertyBag[ConversationSchema.GlobalItemClasses];

    /// <summary>
    ///     Gets the importance of this conversation, calculated by aggregating individual messages importance in the current
    ///     folder only.
    /// </summary>
    public Importance Importance => (Importance)PropertyBag[ConversationSchema.Importance];

    /// <summary>
    ///     Gets the importance of this conversation, calculated by aggregating individual messages importance across all
    ///     folders in the mailbox.
    /// </summary>
    public Importance GlobalImportance => (Importance)PropertyBag[ConversationSchema.GlobalImportance];

    /// <summary>
    ///     Gets the Ids of the messages in this conversation, in the current folder only.
    /// </summary>
    public ItemIdCollection ItemIds => (ItemIdCollection)PropertyBag[ConversationSchema.ItemIds];

    /// <summary>
    ///     Gets the Ids of the messages in this conversation, across all folders in the mailbox.
    /// </summary>
    public ItemIdCollection GlobalItemIds => (ItemIdCollection)PropertyBag[ConversationSchema.GlobalItemIds];

    /// <summary>
    ///     Gets the date and time this conversation was last modified.
    /// </summary>
    public DateTime LastModifiedTime => (DateTime)PropertyBag[ConversationSchema.LastModifiedTime];

    /// <summary>
    ///     Gets the conversation instance key.
    /// </summary>
    public byte[] InstanceKey => (byte[])PropertyBag[ConversationSchema.InstanceKey];

    /// <summary>
    ///     Gets the conversation Preview.
    /// </summary>
    public string Preview => (string)PropertyBag[ConversationSchema.Preview];

    /// <summary>
    ///     Gets the conversation IconIndex.
    /// </summary>
    public IconIndex IconIndex => (IconIndex)PropertyBag[ConversationSchema.IconIndex];

    /// <summary>
    ///     Gets the conversation global IconIndex.
    /// </summary>
    public IconIndex GlobalIconIndex => (IconIndex)PropertyBag[ConversationSchema.GlobalIconIndex];

    /// <summary>
    ///     Gets the draft item ids.
    /// </summary>
    public ItemIdCollection DraftItemIds => (ItemIdCollection)PropertyBag[ConversationSchema.DraftItemIds];

    /// <summary>
    ///     Gets a value indicating if at least one message in this conversation, in the current folder only, is an IRM.
    /// </summary>
    public bool HasIrm => (bool)PropertyBag[ConversationSchema.HasIrm];

    /// <summary>
    ///     Gets a value indicating if at least one message in this conversation, across all folders in the mailbox, is an IRM.
    /// </summary>
    public bool GlobalHasIrm => (bool)PropertyBag[ConversationSchema.GlobalHasIrm];

    #endregion
}
