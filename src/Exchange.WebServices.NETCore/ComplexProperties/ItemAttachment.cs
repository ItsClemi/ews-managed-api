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
///     Represents an item attachment.
/// </summary>
[PublicAPI]
public class ItemAttachment : Attachment
{
    /// <summary>
    ///     The item associated with the attachment.
    /// </summary>
    private Item? _item;

    /// <summary>
    ///     Gets the item associated with the attachment.
    /// </summary>
    public Item? Item
    {
        get => _item;

        internal set
        {
            ThrowIfThisIsNotNew();

            if (_item != null)
            {
                _item.OnChange -= ItemChanged;
            }

            _item = value;

            if (_item != null)
            {
                _item.OnChange += ItemChanged;
            }
        }
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ItemAttachment" /> class.
    /// </summary>
    /// <param name="owner">The owner of the attachment.</param>
    internal ItemAttachment(Item owner)
        : base(owner)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ItemAttachment" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    internal ItemAttachment(ExchangeService service)
        : base(service)
    {
    }

    /// <summary>
    ///     Implements the OnChange event handler for the item associated with the attachment.
    /// </summary>
    /// <param name="serviceObject">The service object that triggered the OnChange event.</param>
    private void ItemChanged(ServiceObject serviceObject)
    {
        if (Owner != null)
        {
            Owner.PropertyBag.Changed();
        }
    }

    /// <summary>
    ///     Obtains EWS XML element name for this object.
    /// </summary>
    /// <returns>The XML element name.</returns>
    internal override string GetXmlElementName()
    {
        return XmlElementNames.ItemAttachment;
    }

    /// <summary>
    ///     Tries to read the element at the current position of the reader.
    /// </summary>
    /// <param name="reader">The reader to read the element from.</param>
    /// <returns>True if the element was read, false otherwise.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        var result = base.TryReadElementFromXml(reader);

        if (!result)
        {
            _item = EwsUtilities.CreateItemFromXmlElementName(this, reader.LocalName);

            if (_item != null)
            {
                _item.LoadFromXml(reader, true);
            }
        }

        return result;
    }

    /// <summary>
    ///     For ItemAttachment, AttachmentId and Item should be patched.
    /// </summary>
    /// <param name="reader"></param>
    /// <returns></returns>
    internal override bool TryReadElementFromXmlToPatch(EwsServiceXmlReader reader)
    {
        // update the attachment id.
        base.TryReadElementFromXml(reader);

        reader.Read();
        var itemClass = EwsUtilities.GetItemTypeFromXmlElementName(reader.LocalName);

        if (itemClass != null)
        {
            if (_item == null || _item.GetType() != itemClass)
            {
                throw new ServiceLocalException(Strings.AttachmentItemTypeMismatch);
            }

            _item.LoadFromXml(reader, false);
            return true;
        }

        return false;
    }

    /// <summary>
    ///     Writes the properties of this object as XML elements.
    /// </summary>
    /// <param name="writer">The writer to write the elements to.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        base.WriteElementsToXml(writer);

        Item.WriteToXml(writer);
    }

    /// <summary>
    ///     Validates this instance.
    /// </summary>
    /// <param name="attachmentIndex">Index of this attachment.</param>
    internal override void Validate(int attachmentIndex)
    {
        if (string.IsNullOrEmpty(Name))
        {
            throw new ServiceValidationException(string.Format(Strings.ItemAttachmentMustBeNamed, attachmentIndex));
        }

        // Recurse through any items attached to item attachment.
        Item.Attachments.Validate();
    }

    /// <summary>
    ///     Loads this attachment.
    /// </summary>
    /// <param name="token"></param>
    /// <param name="additionalProperties">The optional additional properties to load.</param>
    public Task<ServiceResponseCollection<GetAttachmentResponse>> Load(
        CancellationToken token = default,
        params PropertyDefinitionBase[] additionalProperties
    )
    {
        return InternalLoad(null, additionalProperties, token);
    }

    /// <summary>
    ///     Loads this attachment.
    /// </summary>
    /// <param name="additionalProperties">The optional additional properties to load.</param>
    /// <param name="token"></param>
    public Task<ServiceResponseCollection<GetAttachmentResponse>> Load(
        IEnumerable<PropertyDefinitionBase> additionalProperties,
        CancellationToken token = default
    )
    {
        return InternalLoad(null, additionalProperties, token);
    }

    /// <summary>
    ///     Loads this attachment.
    /// </summary>
    /// <param name="bodyType">The body type to load.</param>
    /// <param name="token"></param>
    /// <param name="additionalProperties">The optional additional properties to load.</param>
    public Task<ServiceResponseCollection<GetAttachmentResponse>> Load(
        BodyType bodyType,
        CancellationToken token = default,
        params PropertyDefinitionBase[] additionalProperties
    )
    {
        return InternalLoad(bodyType, additionalProperties, token);
    }

    /// <summary>
    ///     Loads this attachment.
    /// </summary>
    /// <param name="bodyType">The body type to load.</param>
    /// <param name="additionalProperties">The optional additional properties to load.</param>
    /// <param name="token"></param>
    public Task<ServiceResponseCollection<GetAttachmentResponse>> Load(
        BodyType bodyType,
        IEnumerable<PropertyDefinitionBase> additionalProperties,
        CancellationToken token = default
    )
    {
        return InternalLoad(bodyType, additionalProperties, token);
    }
}
