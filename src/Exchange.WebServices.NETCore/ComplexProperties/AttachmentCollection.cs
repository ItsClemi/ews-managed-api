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
using System.Reflection;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents an item's attachment collection.
/// </summary>
[PublicAPI]
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class AttachmentCollection : ComplexPropertyCollection<Attachment>, IOwnedProperty
{
    /// <summary>
    ///     The item owner that owns this attachment collection
    /// </summary>
    private Item _owner;


    /// <summary>
    ///     Initializes a new instance of AttachmentCollection.
    /// </summary>
    internal AttachmentCollection()
    {
    }


    #region IOwnedProperty Members

    /// <summary>
    ///     The owner of this attachment collection.
    /// </summary>
    ServiceObject IOwnedProperty.Owner
    {
        get => _owner;

        set
        {
            var item = value as Item;

            EwsUtilities.Assert(
                item != null,
                "AttachmentCollection.IOwnedProperty.set_Owner",
                "value is not a descendant of ItemBase"
            );

            _owner = item;
        }
    }

    #endregion


    #region Methods

    /// <summary>
    ///     Adds a file attachment to the collection.
    /// </summary>
    /// <param name="fileName">The name of the file representing the content of the attachment.</param>
    /// <returns>A FileAttachment instance.</returns>
    public FileAttachment AddFileAttachment(string fileName)
    {
        return AddFileAttachment(Path.GetFileName(fileName), fileName);
    }

    /// <summary>
    ///     Adds a file attachment to the collection.
    /// </summary>
    /// <param name="name">The display name of the new attachment.</param>
    /// <param name="fileName">The name of the file representing the content of the attachment.</param>
    /// <returns>A FileAttachment instance.</returns>
    public FileAttachment AddFileAttachment(string name, string fileName)
    {
        var fileAttachment = new FileAttachment(_owner)
        {
            Name = name,
            FileName = fileName,
        };

        InternalAdd(fileAttachment);

        return fileAttachment;
    }

    /// <summary>
    ///     Adds a file attachment to the collection.
    /// </summary>
    /// <param name="name">The display name of the new attachment.</param>
    /// <param name="contentStream">The stream from which to read the content of the attachment.</param>
    /// <returns>A FileAttachment instance.</returns>
    public FileAttachment AddFileAttachment(string name, Stream contentStream)
    {
        var fileAttachment = new FileAttachment(_owner)
        {
            Name = name,
            ContentStream = contentStream,
        };

        InternalAdd(fileAttachment);

        return fileAttachment;
    }

    /// <summary>
    ///     Adds a file attachment to the collection.
    /// </summary>
    /// <param name="name">The display name of the new attachment.</param>
    /// <param name="content">A byte arrays representing the content of the attachment.</param>
    /// <returns>A FileAttachment instance.</returns>
    public FileAttachment AddFileAttachment(string name, byte[] content)
    {
        var fileAttachment = new FileAttachment(_owner)
        {
            Name = name,
            Content = content,
        };

        InternalAdd(fileAttachment);

        return fileAttachment;
    }

    /// <summary>
    ///     Adds a reference attachment to the collection
    /// </summary>
    /// <param name="name">The display name of the new attachment.</param>
    /// <param name="attachLongPathName">The fully-qualified path identifying the attachment</param>
    /// <returns>A ReferenceAttachment instance</returns>
    public ReferenceAttachment AddReferenceAttachment(string name, string attachLongPathName)
    {
        var referenceAttachment = new ReferenceAttachment(_owner)
        {
            Name = name,
            AttachLongPathName = attachLongPathName,
        };

        InternalAdd(referenceAttachment);

        return referenceAttachment;
    }

    /// <summary>
    ///     Adds an item attachment to the collection
    /// </summary>
    /// <typeparam name="TItem">The type of the item to attach.</typeparam>
    /// <returns>An ItemAttachment instance.</returns>
    public ItemAttachment<TItem> AddItemAttachment<TItem>()
        where TItem : Item
    {
        if (typeof(TItem).GetTypeInfo().GetCustomAttributes(typeof(AttachableAttribute), false).ToArray().Length == 0)
        {
            throw new InvalidOperationException(
                $"Items of type {typeof(TItem).Name} are not supported as attachments."
            );
        }

        var itemAttachment = new ItemAttachment<TItem>(_owner);
        itemAttachment.Item = (TItem)EwsUtilities.CreateItemFromItemClass(itemAttachment, typeof(TItem), true);

        InternalAdd(itemAttachment);

        return itemAttachment;
    }

    /// <summary>
    ///     Removes all attachments from this collection.
    /// </summary>
    public void Clear()
    {
        InternalClear();
    }

    /// <summary>
    ///     Removes the attachment at the specified index.
    /// </summary>
    /// <param name="index">Index of the attachment to remove.</param>
    public void RemoveAt(int index)
    {
        if (index < 0 || index >= Count)
        {
            throw new ArgumentOutOfRangeException(nameof(index), Strings.IndexIsOutOfRange);
        }

        InternalRemoveAt(index);
    }

    /// <summary>
    ///     Removes the specified attachment.
    /// </summary>
    /// <param name="attachment">The attachment to remove.</param>
    /// <returns>True if the attachment was successfully removed from the collection, false otherwise.</returns>
    public bool Remove(Attachment attachment)
    {
        EwsUtilities.ValidateParam(attachment);

        return InternalRemove(attachment);
    }

    /// <summary>
    ///     Instantiate the appropriate attachment type depending on the current XML element name.
    /// </summary>
    /// <param name="xmlElementName">The XML element name from which to determine the type of attachment to create.</param>
    /// <returns>An Attachment instance.</returns>
    internal override Attachment? CreateComplexProperty(string xmlElementName)
    {
        return xmlElementName switch
        {
            XmlElementNames.FileAttachment => new FileAttachment(_owner),
            XmlElementNames.ItemAttachment => new ItemAttachment(_owner),
            XmlElementNames.ReferenceAttachment => new ReferenceAttachment(_owner),
            _ => null,
        };
    }

    /// <summary>
    ///     Determines the name of the XML element associated with the complexProperty parameter.
    /// </summary>
    /// <param name="complexProperty">The attachment object for which to determine the XML element name with.</param>
    /// <returns>The XML element name associated with the complexProperty parameter.</returns>
    internal override string GetCollectionItemXmlElementName(Attachment complexProperty)
    {
        return complexProperty switch
        {
            FileAttachment => XmlElementNames.FileAttachment,
            ReferenceAttachment => XmlElementNames.ReferenceAttachment,
            _ => XmlElementNames.ItemAttachment,
        };
    }

    /// <summary>
    ///     Saves this collection by creating new attachment and deleting removed ones.
    /// </summary>
    internal async System.Threading.Tasks.Task Save(CancellationToken cancellationToken = default)
    {
        var attachments = new List<Attachment>();

        // Retrieve a list of attachments that have to be deleted.
        foreach (var attachment in RemovedItems)
        {
            if (!attachment.IsNew)
            {
                attachments.Add(attachment);
            }
        }

        // If any, delete them by calling the DeleteAttachment web method.
        if (attachments.Count > 0)
        {
            await InternalDeleteAttachments(attachments, cancellationToken).ConfigureAwait(false);
        }

        attachments.Clear();

        // Retrieve a list of attachments that have to be created.
        foreach (var attachment in this)
        {
            if (attachment.IsNew)
            {
                attachments.Add(attachment);
            }
        }

        // If there are any, create them by calling the CreateAttachment web method.
        if (attachments.Count > 0)
        {
            if (_owner.IsAttachment)
            {
                await InternalCreateAttachments(_owner.ParentAttachment.Id, attachments, cancellationToken)
                    .ConfigureAwait(false);
            }
            else
            {
                await InternalCreateAttachments(_owner.Id.UniqueId, attachments, cancellationToken)
                    .ConfigureAwait(false);
            }
        }

        // Process all of the item attachments in this collection.
        foreach (var attachment in this)
        {
            if (attachment is ItemAttachment itemAttachment)
            {
                // Make sure item was created/loaded before trying to create/delete sub-attachments
                if (itemAttachment.Item != null)
                {
                    // Create/delete any sub-attachments
                    await itemAttachment.Item.Attachments.Save(cancellationToken).ConfigureAwait(false);

                    // Clear the item's change log
                    itemAttachment.Item.ClearChangeLog();
                }
            }
        }

        base.ClearChangeLog();
    }

    /// <summary>
    ///     Determines whether there are any unsaved attachment collection changes.
    /// </summary>
    /// <returns>True if attachment adds or deletes haven't been processed yet.</returns>
    internal bool HasUnprocessedChanges()
    {
        // Any new attachments?
        foreach (var attachment in this)
        {
            if (attachment.IsNew)
            {
                return true;
            }
        }

        // Any pending deletions?
        foreach (var attachment in RemovedItems)
        {
            if (!attachment.IsNew)
            {
                return true;
            }
        }

        // Recurse: process item attachments to check for new or deleted sub-attachments.
        foreach (var itemAttachment in this.OfType<ItemAttachment>())
        {
            if (itemAttachment.Item != null)
            {
                if (itemAttachment.Item.Attachments.HasUnprocessedChanges())
                {
                    return true;
                }
            }
        }

        return false;
    }

    /// <summary>
    ///     Disables the change log clearing mechanism. Attachment collections are saved separately
    ///     from the items they belong to.
    /// </summary>
    internal override void ClearChangeLog()
    {
        // Do nothing
    }

    /// <summary>
    ///     Validates this instance.
    /// </summary>
    internal void Validate()
    {
        // Validate all added attachments
        var contactPhotoFound = false;

        for (var attachmentIndex = 0; attachmentIndex < AddedItems.Count; attachmentIndex++)
        {
            var attachment = AddedItems[attachmentIndex];
            if (attachment.IsNew)
            {
                // At the server side, only the last attachment with IsContactPhoto is kept, all other IsContactPhoto
                // attachments are removed. CreateAttachment will generate AttachmentId for each of such attachments (although
                // only the last one is valid).
                // 
                // With E14 SP2 CreateItemWithAttachment, such request will only return 1 AttachmentId; but the client
                // expects to see all, so let us prevent such "invalid" request in the first place. 
                // 
                // The IsNew check is to still let CreateAttachmentRequest allow multiple IsContactPhoto attachments.
                // 
                if (_owner.IsNew && _owner.Service.RequestedServerVersion >= ExchangeVersion.Exchange2010_SP2)
                {
                    if (attachment is FileAttachment fileAttachment && fileAttachment.IsContactPhoto)
                    {
                        if (contactPhotoFound)
                        {
                            throw new ServiceValidationException(Strings.MultipleContactPhotosInAttachment);
                        }

                        contactPhotoFound = true;
                    }
                }

                attachment.Validate(attachmentIndex);
            }
        }
    }

    /// <summary>
    ///     Calls the DeleteAttachment web method to delete a list of attachments.
    /// </summary>
    /// <param name="attachments">The attachments to delete.</param>
    /// <param name="token"></param>
    private async System.Threading.Tasks.Task InternalDeleteAttachments(
        IEnumerable<Attachment> attachments,
        CancellationToken token
    )
    {
        var responses = await _owner.Service.DeleteAttachments(attachments, token).ConfigureAwait(false);

        foreach (var response in responses)
        {
            // We remove all attachments that were successfully deleted from the change log. We should never
            // receive a warning from EWS, so we ignore them.
            if (response.Result != ServiceResult.Error)
            {
                RemoveFromChangeLog(response.Attachment);
            }
        }

        // TODO : Should we throw for warnings as well?
        if (responses.OverallResult == ServiceResult.Error)
        {
            throw new DeleteAttachmentException(responses, Strings.AtLeastOneAttachmentCouldNotBeDeleted);
        }
    }

    /// <summary>
    ///     Calls the CreateAttachment web method to create a list of attachments.
    /// </summary>
    /// <param name="parentItemId">The Id of the parent item of the new attachments.</param>
    /// <param name="attachments">The attachments to create.</param>
    /// <param name="token"></param>
    private async System.Threading.Tasks.Task InternalCreateAttachments(
        string parentItemId,
        IEnumerable<Attachment> attachments,
        CancellationToken token
    )
    {
        var responses = await _owner.Service.CreateAttachments(parentItemId, attachments, token).ConfigureAwait(false);

        foreach (var response in responses)
        {
            // We remove all attachments that were successfully created from the change log. We should never
            // receive a warning from EWS, so we ignore them.
            if (response.Result != ServiceResult.Error)
            {
                RemoveFromChangeLog(response.Attachment);
            }
        }

        // TODO : Should we throw for warnings as well?
        if (responses.OverallResult == ServiceResult.Error)
        {
            throw new CreateAttachmentException(responses, Strings.AttachmentCreationFailed);
        }
    }

    #endregion
}
