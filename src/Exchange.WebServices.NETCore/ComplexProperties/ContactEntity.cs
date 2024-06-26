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
///     Represents an ContactEntity object.
/// </summary>
[PublicAPI]
public sealed class ContactEntity : ExtractedEntity
{
    /// <summary>
    ///     Gets the contact entity PersonName.
    /// </summary>
    public string PersonName { get; internal set; }

    /// <summary>
    ///     Gets the contact entity BusinessName.
    /// </summary>
    public string BusinessName { get; internal set; }

    /// <summary>
    ///     Gets the contact entity PhoneNumbers.
    /// </summary>
    public ContactPhoneEntityCollection PhoneNumbers { get; internal set; }

    /// <summary>
    ///     Gets the contact entity Urls.
    /// </summary>
    public StringList Urls { get; internal set; }

    /// <summary>
    ///     Gets the contact entity EmailAddresses.
    /// </summary>
    public StringList EmailAddresses { get; internal set; }

    /// <summary>
    ///     Gets the contact entity Addresses.
    /// </summary>
    public StringList Addresses { get; internal set; }

    /// <summary>
    ///     Gets the contact entity ContactString.
    /// </summary>
    public string ContactString { get; internal set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ContactEntity" /> class.
    /// </summary>
    internal ContactEntity()
    {
    }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.NlgPersonName:
            {
                PersonName = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.NlgBusinessName:
            {
                BusinessName = reader.ReadElementValue();
                return true;
            }
            case XmlElementNames.NlgPhoneNumbers:
            {
                PhoneNumbers = new ContactPhoneEntityCollection();
                PhoneNumbers.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.NlgPhoneNumbers);
                return true;
            }
            case XmlElementNames.NlgUrls:
            {
                Urls = new StringList(XmlElementNames.NlgUrl);
                Urls.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.NlgUrls);
                return true;
            }
            case XmlElementNames.NlgEmailAddresses:
            {
                EmailAddresses = new StringList(XmlElementNames.NlgEmailAddress);
                EmailAddresses.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.NlgEmailAddresses);
                return true;
            }
            case XmlElementNames.NlgAddresses:
            {
                Addresses = new StringList(XmlElementNames.NlgAddress);
                Addresses.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.NlgAddresses);
                return true;
            }
            case XmlElementNames.NlgContactString:
            {
                ContactString = reader.ReadElementValue();
                return true;
            }
            default:
            {
                return base.TryReadElementFromXml(reader);
            }
        }
    }
}
