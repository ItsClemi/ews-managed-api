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
///     Represents the base class for e-mail related responses (Reply, Reply all and Forward).
/// </summary>
[PublicAPI]
public sealed class ResponseMessage : ResponseObject<EmailMessage>
{
    /// <summary>
    ///     Gets a value indicating the type of response this object represents.
    /// </summary>
    public ResponseMessageType ResponseType { get; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ResponseMessage" /> class.
    /// </summary>
    /// <param name="referenceItem">The reference item.</param>
    /// <param name="responseType">Type of the response.</param>
    internal ResponseMessage(Item referenceItem, ResponseMessageType responseType)
        : base(referenceItem)
    {
        ResponseType = responseType;
    }

    /// <summary>
    ///     Internal method to return the schema associated with this type of object.
    /// </summary>
    /// <returns>The schema associated with this type of object.</returns>
    internal override ServiceObjectSchema GetSchema()
    {
        return ResponseMessageSchema.Instance;
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
    ///     This methods lets subclasses of ServiceObject override the default mechanism
    ///     by which the XML element name associated with their type is retrieved.
    /// </summary>
    /// <returns>
    ///     The XML element name associated with this type.
    ///     If this method returns null or empty, the XML element name associated with this
    ///     type is determined by the EwsObjectDefinition attribute that decorates the type,
    ///     if present.
    /// </returns>
    /// <remarks>
    ///     Item and folder classes that can be returned by EWS MUST rely on the EwsObjectDefinition
    ///     attribute for XML element name determination.
    /// </remarks>
    internal override string GetXmlElementNameOverride()
    {
        switch (ResponseType)
        {
            case ResponseMessageType.Reply:
            {
                return XmlElementNames.ReplyToItem;
            }
            case ResponseMessageType.ReplyAll:
            {
                return XmlElementNames.ReplyAllToItem;
            }
            case ResponseMessageType.Forward:
            {
                return XmlElementNames.ForwardItem;
            }
            default:
            {
                EwsUtilities.Assert(
                    false,
                    "ResponseMessage.GetXmlElementNameOverride",
                    "An unexpected value for responseType could not be handled."
                );
                return null; // Because the compiler wants it
            }
        }
    }


    #region Properties

    /// <summary>
    ///     Gets or sets the body of the response.
    /// </summary>
    public MessageBody Body
    {
        get => (MessageBody)PropertyBag[ItemSchema.Body];
        set => PropertyBag[ItemSchema.Body] = value;
    }

    /// <summary>
    ///     Gets a list of recipients the response will be sent to.
    /// </summary>
    public EmailAddressCollection ToRecipients => (EmailAddressCollection)PropertyBag[EmailMessageSchema.ToRecipients];

    /// <summary>
    ///     Gets a list of recipients the response will be sent to as Cc.
    /// </summary>
    public EmailAddressCollection CcRecipients => (EmailAddressCollection)PropertyBag[EmailMessageSchema.CcRecipients];

    /// <summary>
    ///     Gets a list of recipients this response will be sent to as Bcc.
    /// </summary>
    public EmailAddressCollection BccRecipients =>
        (EmailAddressCollection)PropertyBag[EmailMessageSchema.BccRecipients];

    /// <summary>
    ///     Gets or sets the subject of this response.
    /// </summary>
    public string Subject
    {
        get => (string)PropertyBag[ItemSchema.Subject];
        set => PropertyBag[ItemSchema.Subject] = value;
    }

    /// <summary>
    ///     Gets or sets the body prefix of this response. The body prefix will be prepended to the original
    ///     message's body when the response is created.
    /// </summary>
    public MessageBody BodyPrefix
    {
        get => (MessageBody)PropertyBag[ResponseObjectSchema.BodyPrefix];
        set => PropertyBag[ResponseObjectSchema.BodyPrefix] = value;
    }

    #endregion
}
