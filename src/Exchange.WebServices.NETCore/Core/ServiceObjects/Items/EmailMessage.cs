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
///     Represents an e-mail message. Properties available on e-mail messages are defined in the EmailMessageSchema class.
/// </summary>
[PublicAPI]
[Attachable]
[ServiceObjectDefinition(XmlElementNames.Message)]
public class EmailMessage : Item
{
    /// <summary>
    ///     Initializes an unsaved local instance of <see cref="EmailMessage" />. To bind to an existing e-mail message, use
    ///     EmailMessage.Bind() instead.
    /// </summary>
    /// <param name="service">The ExchangeService object to which the e-mail message will be bound.</param>
    public EmailMessage(ExchangeService service)
        : base(service)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="EmailMessage" /> class.
    /// </summary>
    /// <param name="parentAttachment">The parent attachment.</param>
    internal EmailMessage(ItemAttachment parentAttachment)
        : base(parentAttachment)
    {
    }

    /// <summary>
    ///     Binds to an existing e-mail message and loads the specified set of properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to use to bind to the e-mail message.</param>
    /// <param name="id">The Id of the e-mail message to bind to.</param>
    /// <param name="propertySet">The set of properties to load.</param>
    /// <param name="token"></param>
    /// <returns>An EmailMessage instance representing the e-mail message corresponding to the specified Id.</returns>
    public new static Task<EmailMessage> Bind(
        ExchangeService service,
        ItemId id,
        PropertySet propertySet,
        CancellationToken token = default
    )
    {
        return service.BindToItem<EmailMessage>(id, propertySet, token);
    }

    /// <summary>
    ///     Binds to an existing e-mail message and loads its first class properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to use to bind to the e-mail message.</param>
    /// <param name="id">The Id of the e-mail message to bind to.</param>
    /// <returns>An EmailMessage instance representing the e-mail message corresponding to the specified Id.</returns>
    public new static Task<EmailMessage> Bind(ExchangeService service, ItemId id)
    {
        return Bind(service, id, PropertySet.FirstClassProperties);
    }

    /// <summary>
    ///     Internal method to return the schema associated with this type of object.
    /// </summary>
    /// <returns>The schema associated with this type of object.</returns>
    internal override ServiceObjectSchema GetSchema()
    {
        return EmailMessageSchema.Instance;
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
    ///     Send message.
    /// </summary>
    /// <param name="parentFolderId">The parent folder id.</param>
    /// <param name="messageDisposition">The message disposition.</param>
    /// <param name="token"></param>
    private async System.Threading.Tasks.Task InternalSend(
        FolderId? parentFolderId,
        MessageDisposition messageDisposition,
        CancellationToken token
    )
    {
        ThrowIfThisIsAttachment();

        if (IsNew)
        {
            if (Attachments.Count == 0 || messageDisposition == MessageDisposition.SaveOnly)
            {
                await InternalCreate(parentFolderId, messageDisposition, null, token).ConfigureAwait(false);
            }
            else
            {
                // If the message has attachments, save as a draft (and add attachments) before sending.
                await InternalCreate(
                        null, // null means use the Drafts folder in the mailbox of the authenticated user.
                        MessageDisposition.SaveOnly,
                        null,
                        token
                    )
                    .ConfigureAwait(false);

                await Service.SendItem(this, parentFolderId, token).ConfigureAwait(false);
            }
        }
        else
        {
            // Regardless of whether item is dirty or not, if it has unprocessed
            // attachment changes, process them now.

            // Validate and save attachments before sending.
            if (HasUnprocessedAttachmentChanges())
            {
                Attachments.Validate();
                await Attachments.Save(token).ConfigureAwait(false);
            }

            if (PropertyBag.GetIsUpdateCallNecessary())
            {
                await InternalUpdate(
                        parentFolderId,
                        ConflictResolutionMode.AutoResolve,
                        messageDisposition,
                        null,
                        token
                    )
                    .ConfigureAwait(false);
            }
            else
            {
                await Service.SendItem(this, parentFolderId, token).ConfigureAwait(false);
            }
        }
    }

    /// <summary>
    ///     Creates a reply response to the message.
    /// </summary>
    /// <param name="replyAll">Indicates whether the reply should go to all of the original recipients of the message.</param>
    /// <returns>A ResponseMessage representing the reply response that can subsequently be modified and sent.</returns>
    public ResponseMessage CreateReply(bool replyAll)
    {
        ThrowIfThisIsNew();

        return new ResponseMessage(this, replyAll ? ResponseMessageType.ReplyAll : ResponseMessageType.Reply);
    }

    /// <summary>
    ///     Creates a forward response to the message.
    /// </summary>
    /// <returns>A ResponseMessage representing the forward response that can subsequently be modified and sent.</returns>
    public ResponseMessage CreateForward()
    {
        ThrowIfThisIsNew();

        return new ResponseMessage(this, ResponseMessageType.Forward);
    }

    /// <summary>
    ///     Replies to the message. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="bodyPrefix">The prefix to prepend to the original body of the message.</param>
    /// <param name="replyAll">Indicates whether the reply should be sent to all of the original recipients of the message.</param>
    public System.Threading.Tasks.Task Reply(MessageBody bodyPrefix, bool replyAll)
    {
        var responseMessage = CreateReply(replyAll);

        responseMessage.BodyPrefix = bodyPrefix;

        return responseMessage.SendAndSaveCopy();
    }

    /// <summary>
    ///     Forwards the message. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="bodyPrefix">The prefix to prepend to the original body of the message.</param>
    /// <param name="toRecipients">The recipients to forward the message to.</param>
    public System.Threading.Tasks.Task Forward(MessageBody bodyPrefix, params EmailAddress[] toRecipients)
    {
        return Forward(bodyPrefix, (IEnumerable<EmailAddress>)toRecipients);
    }

    /// <summary>
    ///     Forwards the message. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="bodyPrefix">The prefix to prepend to the original body of the message.</param>
    /// <param name="toRecipients">The recipients to forward the message to.</param>
    public System.Threading.Tasks.Task Forward(MessageBody bodyPrefix, IEnumerable<EmailAddress> toRecipients)
    {
        var responseMessage = CreateForward();

        responseMessage.BodyPrefix = bodyPrefix;
        responseMessage.ToRecipients.AddRange(toRecipients);

        return responseMessage.SendAndSaveCopy();
    }

    /// <summary>
    ///     Sends this e-mail message. Calling this method results in at least one call to EWS.
    /// </summary>
    public System.Threading.Tasks.Task Send(CancellationToken token = default)
    {
        return InternalSend(null, MessageDisposition.SendOnly, token);
    }

    /// <summary>
    ///     Sends this e-mail message and saves a copy of it in the specified folder. SendAndSaveCopy does not work if the
    ///     message has unsaved attachments. In that case, the message must first be saved and then sent. Calling this method
    ///     results in a call to EWS.
    /// </summary>
    /// <param name="destinationFolderId">The Id of the folder in which to save the copy.</param>
    /// <param name="token"></param>
    public System.Threading.Tasks.Task SendAndSaveCopy(FolderId destinationFolderId, CancellationToken token = default)
    {
        EwsUtilities.ValidateParam(destinationFolderId);

        return InternalSend(destinationFolderId, MessageDisposition.SendAndSaveCopy, token);
    }

    /// <summary>
    ///     Sends this e-mail message and saves a copy of it in the specified folder. SendAndSaveCopy does not work if the
    ///     message has unsaved attachments. In that case, the message must first be saved and then sent. Calling this method
    ///     results in a call to EWS.
    /// </summary>
    /// <param name="destinationFolderName">The name of the folder in which to save the copy.</param>
    /// <param name="token"></param>
    public System.Threading.Tasks.Task SendAndSaveCopy(
        WellKnownFolderName destinationFolderName,
        CancellationToken token = default
    )
    {
        return InternalSend(new FolderId(destinationFolderName), MessageDisposition.SendAndSaveCopy, token);
    }

    /// <summary>
    ///     Sends this e-mail message and saves a copy of it in the Sent Items folder. SendAndSaveCopy does not work if the
    ///     message has unsaved attachments. In that case, the message must first be saved and then sent. Calling this method
    ///     results in a call to EWS.
    /// </summary>
    public System.Threading.Tasks.Task SendAndSaveCopy(CancellationToken token = default)
    {
        return InternalSend(new FolderId(WellKnownFolderName.SentItems), MessageDisposition.SendAndSaveCopy, token);
    }

    /// <summary>
    ///     Suppresses the read receipt on the message. Calling this method results in a call to EWS.
    /// </summary>
    public System.Threading.Tasks.Task SuppressReadReceipt(CancellationToken token = default)
    {
        ThrowIfThisIsNew();

        return new SuppressReadReceipt(this).InternalCreate(null, null, token);
    }


    #region Properties

    /// <summary>
    ///     Gets the list of To recipients for the e-mail message.
    /// </summary>
    public EmailAddressCollection ToRecipients => (EmailAddressCollection)PropertyBag[EmailMessageSchema.ToRecipients];

    /// <summary>
    ///     Gets the list of Bcc recipients for the e-mail message.
    /// </summary>
    public EmailAddressCollection BccRecipients =>
        (EmailAddressCollection)PropertyBag[EmailMessageSchema.BccRecipients];

    /// <summary>
    ///     Gets the Likers associated with the message.
    /// </summary>
    public EmailAddressCollection Likers => (EmailAddressCollection)PropertyBag[EmailMessageSchema.Likers];

    /// <summary>
    ///     Gets the list of Cc recipients for the e-mail message.
    /// </summary>
    public EmailAddressCollection CcRecipients => (EmailAddressCollection)PropertyBag[EmailMessageSchema.CcRecipients];

    /// <summary>
    ///     Gets the conversation topic of the e-mail message.
    /// </summary>
    public string ConversationTopic => (string)PropertyBag[EmailMessageSchema.ConversationTopic];

    /// <summary>
    ///     Gets the conversation index of the e-mail message.
    /// </summary>
    public byte[] ConversationIndex => (byte[])PropertyBag[EmailMessageSchema.ConversationIndex];

    /// <summary>
    ///     Gets or sets the "on behalf" sender of the e-mail message.
    /// </summary>
    public EmailAddress From
    {
        get => (EmailAddress)PropertyBag[EmailMessageSchema.From];
        set => PropertyBag[EmailMessageSchema.From] = value;
    }

    /// <summary>
    ///     Gets or sets a value indicating whether this is an associated message.
    /// </summary>
    public new bool IsAssociated
    {
        get => base.IsAssociated;

        // The "new" keyword is used to expose the setter only on Message types, because
        // EWS only supports creation of FAI Message types.  IsAssociated is a readonly
        // property of the Item type but it is used by the CreateItem web method for creating
        // associated messages.
        set => PropertyBag[ItemSchema.IsAssociated] = value;
    }

    /// <summary>
    ///     Gets or sets a value indicating whether a read receipt is requested for the e-mail message.
    /// </summary>
    public bool IsDeliveryReceiptRequested
    {
        get => (bool)PropertyBag[EmailMessageSchema.IsDeliveryReceiptRequested];
        set => PropertyBag[EmailMessageSchema.IsDeliveryReceiptRequested] = value;
    }

    /// <summary>
    ///     Gets or sets a value indicating whether the e-mail message is read.
    /// </summary>
    public bool IsRead
    {
        get => (bool)PropertyBag[EmailMessageSchema.IsRead];
        set => PropertyBag[EmailMessageSchema.IsRead] = value;
    }

    /// <summary>
    ///     Gets or sets a value indicating whether a read receipt is requested for the e-mail message.
    /// </summary>
    public bool IsReadReceiptRequested
    {
        get => (bool)PropertyBag[EmailMessageSchema.IsReadReceiptRequested];
        set => PropertyBag[EmailMessageSchema.IsReadReceiptRequested] = value;
    }

    /// <summary>
    ///     Gets or sets a value indicating whether a response is requested for the e-mail message.
    /// </summary>
    public bool IsResponseRequested
    {
        get => (bool)PropertyBag[EmailMessageSchema.IsResponseRequested];
        set => PropertyBag[EmailMessageSchema.IsResponseRequested] = value;
    }

    /// <summary>
    ///     Gets the Internet Message Id of the e-mail message.
    /// </summary>
    public string InternetMessageId => (string)PropertyBag[EmailMessageSchema.InternetMessageId];

    /// <summary>
    ///     Gets or sets the references of the e-mail message.
    /// </summary>
    public string References
    {
        get => (string)PropertyBag[EmailMessageSchema.References];
        set => PropertyBag[EmailMessageSchema.References] = value;
    }

    /// <summary>
    ///     Gets a list of e-mail addresses to which replies should be addressed.
    /// </summary>
    public EmailAddressCollection ReplyTo => (EmailAddressCollection)PropertyBag[EmailMessageSchema.ReplyTo];

    /// <summary>
    ///     Gets or sets the sender of the e-mail message.
    /// </summary>
    public EmailAddress Sender
    {
        get => (EmailAddress)PropertyBag[EmailMessageSchema.Sender];
        set => PropertyBag[EmailMessageSchema.Sender] = value;
    }

    /// <summary>
    ///     Gets the ReceivedBy property of the e-mail message.
    /// </summary>
    public EmailAddress ReceivedBy => (EmailAddress)PropertyBag[EmailMessageSchema.ReceivedBy];

    /// <summary>
    ///     Gets the ReceivedRepresenting property of the e-mail message.
    /// </summary>
    public EmailAddress ReceivedRepresenting => (EmailAddress)PropertyBag[EmailMessageSchema.ReceivedRepresenting];

    /// <summary>
    ///     Gets the ApprovalRequestData property of the e-mail message.
    /// </summary>
    public ApprovalRequestData ApprovalRequestData =>
        (ApprovalRequestData)PropertyBag[EmailMessageSchema.ApprovalRequestData];

    /// <summary>
    ///     Gets the VotingInformation property of the e-mail message.
    /// </summary>
    public VotingInformation VotingInformation => (VotingInformation)PropertyBag[EmailMessageSchema.VotingInformation];

    #endregion
}
