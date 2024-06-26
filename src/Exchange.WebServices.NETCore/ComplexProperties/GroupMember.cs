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
///     Represents a group member.
/// </summary>
[PublicAPI]
[RequiredServerVersion(ExchangeVersion.Exchange2010)]
public class GroupMember : ComplexProperty
{
    /// <summary>
    ///     AddressInformation field.
    /// </summary>
    private EmailAddress? _addressInformation;

    /// <summary>
    ///     Gets the key of the member.
    /// </summary>
    public string? Key { get; private set; }

    /// <summary>
    ///     Gets the address information of the member.
    /// </summary>
    public EmailAddress? AddressInformation
    {
        get => _addressInformation;

        internal set
        {
            if (_addressInformation != null)
            {
                _addressInformation.OnChange -= AddressInformationChanged;
            }

            _addressInformation = value;

            if (_addressInformation != null)
            {
                _addressInformation.OnChange += AddressInformationChanged;
            }
        }
    }

    /// <summary>
    ///     Gets the status of the member.
    /// </summary>
    public MemberStatus Status { get; private set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="GroupMember" /> class.
    /// </summary>
    public GroupMember()
    {
        // Key is assigned by server
        Key = null;

        // Member status is calculated by server
        Status = MemberStatus.Unrecognized;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="GroupMember" /> class.
    /// </summary>
    /// <param name="smtpAddress">The SMTP address of the member.</param>
    public GroupMember(string smtpAddress)
        : this()
    {
        AddressInformation = new EmailAddress(smtpAddress);
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="GroupMember" /> class.
    /// </summary>
    /// <param name="address">The address of the member.</param>
    /// <param name="routingType">The routing type of the address.</param>
    /// <param name="mailboxType">The mailbox type of the member.</param>
    public GroupMember(string address, string routingType, MailboxType mailboxType)
        : this()
    {
        switch (mailboxType)
        {
            case MailboxType.PublicGroup:
            case MailboxType.PublicFolder:
            case MailboxType.Mailbox:
            case MailboxType.Contact:
            case MailboxType.OneOff:
            {
                AddressInformation = new EmailAddress(null, address, routingType, mailboxType);
                break;
            }
            default:
            {
                throw new ServiceLocalException(Strings.InvalidMailboxType);
            }
        }
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="GroupMember" /> class.
    /// </summary>
    /// <param name="smtpAddress">The SMTP address of the member.</param>
    /// <param name="mailboxType">The mailbox type of the member.</param>
    public GroupMember(string smtpAddress, MailboxType mailboxType)
        : this(smtpAddress, EmailAddress.SmtpRoutingType, mailboxType)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="GroupMember" /> class.
    /// </summary>
    /// <param name="name">The name of the one-off member.</param>
    /// <param name="address">The address of the one-off member.</param>
    /// <param name="routingType">The routing type of the address.</param>
    public GroupMember(string name, string address, string routingType = EmailAddress.SmtpRoutingType)
        : this()
    {
        AddressInformation = new EmailAddress(name, address, routingType, MailboxType.OneOff);
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="GroupMember" /> class.
    /// </summary>
    /// <param name="contactGroupId">The Id of the contact group to link the member to.</param>
    public GroupMember(ItemId contactGroupId)
        : this()
    {
        AddressInformation = new EmailAddress(null, null, null, MailboxType.ContactGroup, contactGroupId);
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="GroupMember" /> class.
    /// </summary>
    /// <param name="contactId">The Id of the contact member.</param>
    /// <param name="addressToLink">The Id of the contact to link the member to.</param>
    public GroupMember(ItemId contactId, string? addressToLink)
        : this()
    {
        AddressInformation = new EmailAddress(null, addressToLink, null, MailboxType.Contact, contactId);
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="GroupMember" /> class.
    /// </summary>
    /// <param name="addressInformation">The e-mail address of the member.</param>
    public GroupMember(EmailAddress addressInformation)
        : this()
    {
        AddressInformation = new EmailAddress(addressInformation);
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="GroupMember" /> class from another GroupMember instance.
    /// </summary>
    /// <param name="member">GroupMember class instance to copy.</param>
    internal GroupMember(GroupMember member)
        : this()
    {
        EwsUtilities.ValidateParam(member);
        AddressInformation = new EmailAddress(member.AddressInformation);
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="GroupMember" /> class from a Contact instance indexed by the specified
    ///     key.
    /// </summary>
    /// <param name="contact">The contact to link to.</param>
    /// <param name="emailAddressKey">The contact's e-mail address to link to.</param>
    public GroupMember(Contact contact, EmailAddressKey emailAddressKey)
        : this()
    {
        EwsUtilities.ValidateParam(contact);

        var emailAddress = contact.EmailAddresses[emailAddressKey];

        AddressInformation = new EmailAddress(emailAddress);

        _addressInformation.Id = contact.Id;
    }

    /// <summary>
    ///     Reads the member Key attribute from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
    {
        Key = reader.ReadAttributeValue<string>(XmlAttributeNames.Key);
    }

    /// <summary>
    ///     Tries to read Status or Mailbox elements from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.Status:
            {
                Status = EwsUtilities.Parse<MemberStatus>(reader.ReadElementValue());
                return true;
            }
            case XmlElementNames.Mailbox:
            {
                AddressInformation = new EmailAddress();
                AddressInformation.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            default:
            {
                return false;
            }
        }
    }

    /// <summary>
    ///     Writes the member key attribute to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        // if this.key is null or empty, writer skips the attribute
        writer.WriteAttributeValue(XmlAttributeNames.Key, Key);
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        // No need to write member Status back to server
        // Write only AddressInformation container element
        AddressInformation.WriteToXml(writer, XmlNamespace.Types, XmlElementNames.Mailbox);
    }

    /// <summary>
    ///     AddressInformation instance is changed.
    /// </summary>
    /// <param name="complexProperty">Changed property.</param>
    private void AddressInformationChanged(ComplexProperty complexProperty)
    {
        Changed();
    }
}
