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
///     Represents a collection of e-mail addresses.
/// </summary>
[PublicAPI]
public sealed class EmailAddressCollection : ComplexPropertyCollection<EmailAddress>
{
    /// <summary>
    ///     XML element name
    /// </summary>
    private readonly string _collectionItemXmlElementName;

    /// <summary>
    ///     Initializes a new instance of the <see cref="EmailAddressCollection" /> class.
    /// </summary>
    /// <remarks>
    ///     Note that XmlElementNames.Mailbox is the collection element name for ArrayOfRecipientsType, not
    ///     ArrayOfEmailAddressesType.
    /// </remarks>
    internal EmailAddressCollection()
        : this(XmlElementNames.Mailbox)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="EmailAddressCollection" /> class.
    /// </summary>
    /// <param name="collectionItemXmlElementName">Name of the collection item XML element.</param>
    internal EmailAddressCollection(string collectionItemXmlElementName)
    {
        _collectionItemXmlElementName = collectionItemXmlElementName;
    }

    /// <summary>
    ///     Adds an e-mail address to the collection.
    /// </summary>
    /// <param name="emailAddress">The e-mail address to add.</param>
    public void Add(EmailAddress emailAddress)
    {
        InternalAdd(emailAddress);
    }

    /// <summary>
    ///     Adds multiple e-mail addresses to the collection.
    /// </summary>
    /// <param name="emailAddresses">The e-mail addresses to add.</param>
    public void AddRange(IEnumerable<EmailAddress> emailAddresses)
    {
        foreach (var emailAddress in emailAddresses)
        {
            Add(emailAddress);
        }
    }

    /// <summary>
    ///     Adds an e-mail address to the collection.
    /// </summary>
    /// <param name="smtpAddress">The SMTP address used to initialize the e-mail address.</param>
    /// <returns>An EmailAddress object initialized with the provided SMTP address.</returns>
    public EmailAddress Add(string smtpAddress)
    {
        var emailAddress = new EmailAddress(smtpAddress);

        Add(emailAddress);

        return emailAddress;
    }

    /// <summary>
    ///     Adds multiple e-mail addresses to the collection.
    /// </summary>
    /// <param name="smtpAddresses">The SMTP addresses used to initialize the e-mail addresses.</param>
    public void AddRange(IEnumerable<string> smtpAddresses)
    {
        foreach (var smtpAddress in smtpAddresses)
        {
            Add(smtpAddress);
        }
    }

    /// <summary>
    ///     Adds an e-mail address to the collection.
    /// </summary>
    /// <param name="name">The name used to initialize the e-mail address.</param>
    /// <param name="smtpAddress">The SMTP address used to initialize the e-mail address.</param>
    /// <returns>An EmailAddress object initialized with the provided SMTP address.</returns>
    public EmailAddress Add(string name, string smtpAddress)
    {
        var emailAddress = new EmailAddress(name, smtpAddress);

        Add(emailAddress);

        return emailAddress;
    }

    /// <summary>
    ///     Clears the collection.
    /// </summary>
    public void Clear()
    {
        InternalClear();
    }

    /// <summary>
    ///     Removes an e-mail address from the collection.
    /// </summary>
    /// <param name="index">The index of the e-mail address to remove.</param>
    public void RemoveAt(int index)
    {
        if (index < 0 || index >= Count)
        {
            throw new ArgumentOutOfRangeException(nameof(index), Strings.IndexIsOutOfRange);
        }

        InternalRemoveAt(index);
    }

    /// <summary>
    ///     Removes an e-mail address from the collection.
    /// </summary>
    /// <param name="emailAddress">The e-mail address to remove.</param>
    /// <returns>True if the email address was successfully removed from the collection, false otherwise.</returns>
    public bool Remove(EmailAddress emailAddress)
    {
        EwsUtilities.ValidateParam(emailAddress);

        return InternalRemove(emailAddress);
    }

    /// <summary>
    ///     Creates an EmailAddress object from an XML element name.
    /// </summary>
    /// <param name="xmlElementName">The XML element name from which to create the e-mail address.</param>
    /// <returns>An EmailAddress object.</returns>
    internal override EmailAddress? CreateComplexProperty(string xmlElementName)
    {
        if (xmlElementName == _collectionItemXmlElementName)
        {
            return new EmailAddress();
        }

        return null;
    }

    /// <summary>
    ///     Retrieves the XML element name corresponding to the provided EmailAddress object.
    /// </summary>
    /// <param name="emailAddress">The EmailAddress object from which to determine the XML element name.</param>
    /// <returns>The XML element name corresponding to the provided EmailAddress object.</returns>
    internal override string GetCollectionItemXmlElementName(EmailAddress emailAddress)
    {
        return _collectionItemXmlElementName;
    }

    /// <summary>
    ///     Determine whether we should write collection to XML or not.
    /// </summary>
    /// <returns>Always true, even if the collection is empty.</returns>
    internal override bool ShouldWriteToRequest()
    {
        return true;
    }
}
