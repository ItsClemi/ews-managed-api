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

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents an entry of an EmailAddressDictionary.
/// </summary>
[PublicAPI]
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class EmailAddressEntry : DictionaryEntryProperty<EmailAddressKey>
{
    /// <summary>
    ///     The email address.
    /// </summary>
    private EmailAddress? _emailAddress;

    /// <summary>
    ///     Gets or sets the e-mail address of the entry.
    /// </summary>
    public EmailAddress? EmailAddress
    {
        get => _emailAddress;

        set
        {
            SetFieldValue(ref _emailAddress, value);

            if (_emailAddress != null)
            {
                _emailAddress.OnChange += EmailAddressChanged;
            }
        }
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="EmailAddressEntry" /> class.
    /// </summary>
    internal EmailAddressEntry()
    {
        _emailAddress = new EmailAddress();
        _emailAddress.OnChange += EmailAddressChanged;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="EmailAddressEntry" /> class.
    /// </summary>
    /// <param name="key">The key.</param>
    /// <param name="emailAddress">The email address.</param>
    internal EmailAddressEntry(EmailAddressKey key, EmailAddress? emailAddress)
        : base(key)
    {
        _emailAddress = emailAddress;

        if (_emailAddress != null)
        {
            _emailAddress.OnChange += EmailAddressChanged;
        }
    }

    /// <summary>
    ///     Reads the attributes from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
    {
        base.ReadAttributesFromXml(reader);

        EmailAddress.Name = reader.ReadAttributeValue<string>(XmlAttributeNames.Name);
        EmailAddress.RoutingType = reader.ReadAttributeValue<string>(XmlAttributeNames.RoutingType);

        var mailboxTypeString = reader.ReadAttributeValue(XmlAttributeNames.MailboxType);
        if (!string.IsNullOrEmpty(mailboxTypeString))
        {
            EmailAddress.MailboxType = EwsUtilities.Parse<MailboxType>(mailboxTypeString);
        }
        else
        {
            EmailAddress.MailboxType = null;
        }
    }

    /// <summary>
    ///     Reads the text value from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadTextValueFromXml(EwsServiceXmlReader reader)
    {
        EmailAddress.Address = reader.ReadValue();
    }

    /// <summary>
    ///     Writes the attributes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        base.WriteAttributesToXml(writer);

        if (writer.Service.RequestedServerVersion > ExchangeVersion.Exchange2007_SP1)
        {
            writer.WriteAttributeValue(XmlAttributeNames.Name, EmailAddress.Name);
            writer.WriteAttributeValue(XmlAttributeNames.RoutingType, EmailAddress.RoutingType);
            if (EmailAddress.MailboxType != MailboxType.Unknown)
            {
                writer.WriteAttributeValue(XmlAttributeNames.MailboxType, EmailAddress.MailboxType);
            }
        }
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteValue(EmailAddress.Address, XmlElementNames.EmailAddress);
    }

    /// <summary>
    ///     E-mail address was changed.
    /// </summary>
    /// <param name="complexProperty">Property that changed.</param>
    private void EmailAddressChanged(ComplexProperty complexProperty)
    {
        Changed();
    }
}
