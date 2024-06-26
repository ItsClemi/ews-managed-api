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
///     Represents an abstract CreateItem request.
/// </summary>
/// <typeparam name="TServiceObject">The type of the service object.</typeparam>
/// <typeparam name="TResponse">The type of the response.</typeparam>
internal abstract class CreateItemRequestBase<TServiceObject, TResponse> : CreateRequest<TServiceObject, TResponse>
    where TServiceObject : ServiceObject
    where TResponse : ServiceResponse
{
    /// <summary>
    ///     Gets a value indicating whether the TimeZoneContext SOAP header should be emitted.
    /// </summary>
    /// <value>
    ///     <c>true</c> if the time zone should be emitted; otherwise, <c>false</c>.
    /// </value>
    internal override bool EmitTimeZoneHeader
    {
        get
        {
            foreach (var serviceObject in Items)
            {
                if (serviceObject.GetIsTimeZoneHeaderRequired(false))
                {
                    return true;
                }
            }

            return false;
        }
    }

    /// <summary>
    ///     Gets or sets the message disposition.
    /// </summary>
    /// <value>The message disposition.</value>
    public MessageDisposition? MessageDisposition { get; set; }

    /// <summary>
    ///     Gets or sets the send invitations mode.
    /// </summary>
    /// <value>The send invitations mode.</value>
    public SendInvitationsMode? SendInvitationsMode { get; set; }

    /// <summary>
    ///     Gets or sets the items.
    /// </summary>
    /// <value>The items.</value>
    public IEnumerable<TServiceObject> Items
    {
        get => Objects;
        set => Objects = value;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="CreateItemRequestBase&lt;TServiceObject, TResponse&gt;" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
    protected CreateItemRequestBase(ExchangeService service, ServiceErrorHandling errorHandlingMode)
        : base(service, errorHandlingMode)
    {
    }

    /// <summary>
    ///     Validate the request.
    /// </summary>
    internal override void Validate()
    {
        base.Validate();

        EwsUtilities.ValidateParam(Items);
    }

    /// <summary>
    ///     Gets the name of the XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal override string GetXmlElementName()
    {
        return XmlElementNames.CreateItem;
    }

    /// <summary>
    ///     Gets the name of the response XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal override string GetResponseXmlElementName()
    {
        return XmlElementNames.CreateItemResponse;
    }

    /// <summary>
    ///     Gets the name of the response message XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    protected override string GetResponseMessageXmlElementName()
    {
        return XmlElementNames.CreateItemResponseMessage;
    }

    /// <summary>
    ///     Gets the name of the parent folder XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal override string GetParentFolderXmlElementName()
    {
        return XmlElementNames.SavedItemFolderId;
    }

    /// <summary>
    ///     Gets the name of the object collection XML element.
    /// </summary>
    /// <returns>XML element name.</returns>
    internal override string GetObjectCollectionXmlElementName()
    {
        return XmlElementNames.Items;
    }

    /// <summary>
    ///     Writes the attributes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        base.WriteAttributesToXml(writer);

        if (MessageDisposition.HasValue)
        {
            writer.WriteAttributeValue(XmlAttributeNames.MessageDisposition, MessageDisposition.Value);
        }

        if (SendInvitationsMode.HasValue)
        {
            writer.WriteAttributeValue(XmlAttributeNames.SendMeetingInvitations, SendInvitationsMode.Value);
        }
    }
}
