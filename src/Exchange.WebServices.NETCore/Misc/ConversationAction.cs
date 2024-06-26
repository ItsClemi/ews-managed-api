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
///     ConversationAction class that represents ConversationActionType in the request XML.
///     This class really is meant for representing single ConversationAction that needs to
///     be taken on a conversation.
/// </summary>
internal class ConversationAction
{
    /// <summary>
    ///     Gets or sets conversation action
    /// </summary>
    internal ConversationActionType Action { get; set; }

    /// <summary>
    ///     Gets or sets conversation id
    /// </summary>
    internal ConversationId ConversationId { get; set; }

    /// <summary>
    ///     Gets or sets ProcessRightAway
    /// </summary>
    internal bool ProcessRightAway { get; set; }

    /// <summary>
    ///     Gets or set conversation categories for Always Categorize action
    /// </summary>
    internal StringList? Categories { get; set; }

    /// <summary>
    ///     Gets or sets Enable Always Delete value for Always Delete action
    /// </summary>
    internal bool EnableAlwaysDelete { get; set; }

    /// <summary>
    ///     Gets or sets the IsRead state.
    /// </summary>
    internal bool? IsRead { get; set; }

    /// <summary>
    ///     Gets or sets the SuppressReadReceipts flag.
    /// </summary>
    internal bool? SuppressReadReceipts { get; set; }

    /// <summary>
    ///     Gets or sets the Deletion mode.
    /// </summary>
    internal DeleteMode? DeleteType { get; set; }

    /// <summary>
    ///     Gets or sets the flag.
    /// </summary>
    internal Flag? Flag { get; set; }

    /// <summary>
    ///     ConversationLastSyncTime is used in one time action to determine the items
    ///     on which to take the action.
    /// </summary>
    internal DateTime? ConversationLastSyncTime { get; set; }

    /// <summary>
    ///     Gets or sets folder id ContextFolder
    /// </summary>
    internal FolderIdWrapper? ContextFolderId { get; set; }

    /// <summary>
    ///     Gets or sets folder id for Move action
    /// </summary>
    internal FolderIdWrapper? DestinationFolderId { get; set; }

    /// <summary>
    ///     Gets or sets the retention policy type.
    /// </summary>
    internal RetentionType? RetentionPolicyType { get; set; }

    /// <summary>
    ///     Gets or sets the retention policy tag id.
    /// </summary>
    internal Guid? RetentionPolicyTagId { get; set; }

    /// <summary>
    ///     Gets the name of the XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal string GetXmlElementName()
    {
        return XmlElementNames.ApplyConversationAction;
    }

    /// <summary>
    ///     Validate request.
    /// </summary>
    internal void Validate()
    {
        EwsUtilities.ValidateParam(ConversationId, "conversationId");
    }

    /// <summary>
    ///     Writes XML elements.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.ConversationAction);
        try
        {
            var actionValue = Action switch
            {
                ConversationActionType.AlwaysCategorize => XmlElementNames.AlwaysCategorize,
                ConversationActionType.AlwaysDelete => XmlElementNames.AlwaysDelete,
                ConversationActionType.AlwaysMove => XmlElementNames.AlwaysMove,
                ConversationActionType.Delete => XmlElementNames.Delete,
                ConversationActionType.Copy => XmlElementNames.Copy,
                ConversationActionType.Move => XmlElementNames.Move,
                ConversationActionType.SetReadState => XmlElementNames.SetReadState,
                ConversationActionType.SetRetentionPolicy => XmlElementNames.SetRetentionPolicy,
                ConversationActionType.Flag => XmlElementNames.Flag,
                _ => throw new ArgumentException("ConversationAction"),
            };

            // Emit the action element
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Action, actionValue);

            // Emit the conversation id element
            ConversationId.WriteToXml(writer, XmlNamespace.Types, XmlElementNames.ConversationId);

            if (Action == ConversationActionType.AlwaysCategorize ||
                Action == ConversationActionType.AlwaysDelete ||
                Action == ConversationActionType.AlwaysMove)
            {
                // Emit the ProcessRightAway element
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.ProcessRightAway,
                    EwsUtilities.BoolToXsBool(ProcessRightAway)
                );
            }

            if (Action == ConversationActionType.AlwaysCategorize)
            {
                // Emit the categories element
                if (Categories != null && Categories.Count > 0)
                {
                    Categories.WriteToXml(writer, XmlNamespace.Types, XmlElementNames.Categories);
                }
            }
            else if (Action == ConversationActionType.AlwaysDelete)
            {
                // Emit the EnableAlwaysDelete element
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.EnableAlwaysDelete,
                    EwsUtilities.BoolToXsBool(EnableAlwaysDelete)
                );
            }
            else if (Action == ConversationActionType.AlwaysMove)
            {
                // Emit the Move Folder Id
                if (DestinationFolderId != null)
                {
                    writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.DestinationFolderId);
                    DestinationFolderId.WriteToXml(writer);
                    writer.WriteEndElement();
                }
            }
            else
            {
                if (ContextFolderId != null)
                {
                    writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.ContextFolderId);

                    ContextFolderId.WriteToXml(writer);

                    writer.WriteEndElement();
                }

                if (ConversationLastSyncTime.HasValue)
                {
                    writer.WriteElementValue(
                        XmlNamespace.Types,
                        XmlElementNames.ConversationLastSyncTime,
                        ConversationLastSyncTime.Value
                    );
                }

                if (Action == ConversationActionType.Copy)
                {
                    EwsUtilities.Assert(
                        DestinationFolderId != null,
                        "ApplyconversationActionRequest",
                        "DestinationFolderId should be set when performing copy action"
                    );

                    writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.DestinationFolderId);
                    DestinationFolderId.WriteToXml(writer);
                    writer.WriteEndElement();
                }
                else if (Action == ConversationActionType.Move)
                {
                    EwsUtilities.Assert(
                        DestinationFolderId != null,
                        "ApplyconversationActionRequest",
                        "DestinationFolderId should be set when performing move action"
                    );

                    writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.DestinationFolderId);
                    DestinationFolderId.WriteToXml(writer);
                    writer.WriteEndElement();
                }
                else if (Action == ConversationActionType.Delete)
                {
                    EwsUtilities.Assert(
                        DeleteType.HasValue,
                        "ApplyconversationActionRequest",
                        "DeleteType should be specified when deleting a conversation."
                    );

                    writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.DeleteType, DeleteType.Value);
                }
                else if (Action == ConversationActionType.SetReadState)
                {
                    EwsUtilities.Assert(
                        IsRead.HasValue,
                        "ApplyconversationActionRequest",
                        "IsRead should be specified when marking/unmarking a conversation as read."
                    );

                    writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.IsRead, IsRead.Value);

                    if (SuppressReadReceipts.HasValue)
                    {
                        writer.WriteElementValue(
                            XmlNamespace.Types,
                            XmlElementNames.SuppressReadReceipts,
                            SuppressReadReceipts.Value
                        );
                    }
                }
                else if (Action == ConversationActionType.SetRetentionPolicy)
                {
                    EwsUtilities.Assert(
                        RetentionPolicyType.HasValue,
                        "ApplyconversationActionRequest",
                        "RetentionPolicyType should be specified when setting a retention policy on a conversation."
                    );

                    writer.WriteElementValue(
                        XmlNamespace.Types,
                        XmlElementNames.RetentionPolicyType,
                        RetentionPolicyType.Value
                    );

                    if (RetentionPolicyTagId.HasValue)
                    {
                        writer.WriteElementValue(
                            XmlNamespace.Types,
                            XmlElementNames.RetentionPolicyTagId,
                            RetentionPolicyTagId.Value
                        );
                    }
                }
                else if (Action == ConversationActionType.Flag)
                {
                    EwsUtilities.Assert(
                        Flag != null,
                        "ApplyconversationActionRequest",
                        "Flag should be specified when flagging conversation items."
                    );

                    writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Flag);
                    Flag.WriteElementsToXml(writer);
                    writer.WriteEndElement();
                }
            }
        }
        finally
        {
            writer.WriteEndElement();
        }
    }
}
