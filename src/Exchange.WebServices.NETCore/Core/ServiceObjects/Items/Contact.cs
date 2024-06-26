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
///     Represents a contact. Properties available on contacts are defined in the ContactSchema class.
/// </summary>
[PublicAPI]
[Attachable]
[ServiceObjectDefinition(XmlElementNames.Contact)]
public class Contact : Item
{
    private const string ContactPictureName = "ContactPicture.jpg";

    /// <summary>
    ///     Initializes an unsaved local instance of <see cref="Contact" />. To bind to an existing contact, use Contact.Bind()
    ///     instead.
    /// </summary>
    /// <param name="service">The ExchangeService object to which the contact will be bound.</param>
    public Contact(ExchangeService service)
        : base(service)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="Contact" /> class.
    /// </summary>
    /// <param name="parentAttachment">The parent attachment.</param>
    internal Contact(ItemAttachment parentAttachment)
        : base(parentAttachment)
    {
    }

    /// <summary>
    ///     Binds to an existing contact and loads the specified set of properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to use to bind to the contact.</param>
    /// <param name="id">The Id of the contact to bind to.</param>
    /// <param name="propertySet">The set of properties to load.</param>
    /// <param name="token"></param>
    /// <returns>A Contact instance representing the contact corresponding to the specified Id.</returns>
    public new static Task<Contact> Bind(
        ExchangeService service,
        ItemId id,
        PropertySet propertySet,
        CancellationToken token = default
    )
    {
        return service.BindToItem<Contact>(id, propertySet, token);
    }

    /// <summary>
    ///     Binds to an existing contact and loads its first class properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to use to bind to the contact.</param>
    /// <param name="id">The Id of the contact to bind to.</param>
    /// <returns>A Contact instance representing the contact corresponding to the specified Id.</returns>
    public new static Task<Contact> Bind(ExchangeService service, ItemId id)
    {
        return Bind(service, id, PropertySet.FirstClassProperties);
    }

    /// <summary>
    ///     Internal method to return the schema associated with this type of object.
    /// </summary>
    /// <returns>The schema associated with this type of object.</returns>
    internal override ServiceObjectSchema GetSchema()
    {
        return ContactSchema.Instance;
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
    ///     Sets the contact's picture using the specified byte array.
    /// </summary>
    /// <param name="content">The bytes making up the picture.</param>
    public void SetContactPicture(byte[] content)
    {
        EwsUtilities.ValidateMethodVersion(Service, ExchangeVersion.Exchange2010, "SetContactPicture");

        InternalRemoveContactPicture();
        var fileAttachment = Attachments.AddFileAttachment(ContactPictureName, content);
        fileAttachment.IsContactPhoto = true;
    }

    /// <summary>
    ///     Sets the contact's picture using the specified stream.
    /// </summary>
    /// <param name="contentStream">The stream containing the picture.</param>
    public void SetContactPicture(Stream contentStream)
    {
        EwsUtilities.ValidateMethodVersion(Service, ExchangeVersion.Exchange2010, "SetContactPicture");

        InternalRemoveContactPicture();
        var fileAttachment = Attachments.AddFileAttachment(ContactPictureName, contentStream);
        fileAttachment.IsContactPhoto = true;
    }

    /// <summary>
    ///     Sets the contact's picture using the specified file.
    /// </summary>
    /// <param name="fileName">The name of the file that contains the picture.</param>
    public void SetContactPicture(string fileName)
    {
        EwsUtilities.ValidateMethodVersion(Service, ExchangeVersion.Exchange2010, "SetContactPicture");

        InternalRemoveContactPicture();
        var fileAttachment = Attachments.AddFileAttachment(Path.GetFileName(fileName), fileName);
        fileAttachment.IsContactPhoto = true;
    }

    /// <summary>
    ///     Retrieves the file attachment that holds the contact's picture.
    /// </summary>
    /// <returns>The file attachment that holds the contact's picture.</returns>
    public FileAttachment? GetContactPictureAttachment()
    {
        EwsUtilities.ValidateMethodVersion(Service, ExchangeVersion.Exchange2010, "GetContactPictureAttachment");

        if (!PropertyBag.IsPropertyLoaded(ItemSchema.Attachments))
        {
            throw new PropertyException(Strings.AttachmentCollectionNotLoaded);
        }

        foreach (var attachment in Attachments)
        {
            var fileAttachment = (FileAttachment)attachment;
            if (fileAttachment.IsContactPhoto)
            {
                return fileAttachment;
            }
        }

        return null;
    }

    /// <summary>
    ///     Removes the picture from local attachment collection.
    /// </summary>
    private void InternalRemoveContactPicture()
    {
        // Iterates in reverse order to remove file attachments that have IsContactPhoto set to true.
        for (var index = Attachments.Count - 1; index >= 0; index--)
        {
            if (Attachments[index] is FileAttachment fileAttachment)
            {
                if (fileAttachment.IsContactPhoto)
                {
                    Attachments.Remove(fileAttachment);
                }
            }
        }
    }

    /// <summary>
    ///     Removes the contact's picture.
    /// </summary>
    public void RemoveContactPicture()
    {
        EwsUtilities.ValidateMethodVersion(Service, ExchangeVersion.Exchange2010, "RemoveContactPicture");

        if (!PropertyBag.IsPropertyLoaded(ItemSchema.Attachments))
        {
            throw new PropertyException(Strings.AttachmentCollectionNotLoaded);
        }

        InternalRemoveContactPicture();
    }

    /// <summary>
    ///     Validates this instance.
    /// </summary>
    internal override void Validate()
    {
        base.Validate();

        if (TryGetProperty(ContactSchema.FileAsMapping, out var fileAsMapping))
        {
            // FileAsMapping is extended by 5 new values in 2010 mode. Validate that they are used according the version.
            EwsUtilities.ValidateEnumVersionValue((FileAsMapping)fileAsMapping, Service.RequestedServerVersion);
        }
    }


    #region Properties

    /// <summary>
    ///     Gets or set the name under which this contact is filed as. FileAs can be manually set or
    ///     can be automatically calculated based on the value of the FileAsMapping property.
    /// </summary>
    public string FileAs
    {
        get => (string)PropertyBag[ContactSchema.FileAs];
        set => PropertyBag[ContactSchema.FileAs] = value;
    }

    /// <summary>
    ///     Gets or sets a value indicating how the FileAs property should be automatically calculated.
    /// </summary>
    public FileAsMapping FileAsMapping
    {
        get => (FileAsMapping)PropertyBag[ContactSchema.FileAsMapping];
        set => PropertyBag[ContactSchema.FileAsMapping] = value;
    }

    /// <summary>
    ///     Gets or sets the display name of the contact.
    /// </summary>
    public string DisplayName
    {
        get => (string)PropertyBag[ContactSchema.DisplayName];
        set => PropertyBag[ContactSchema.DisplayName] = value;
    }

    /// <summary>
    ///     Gets or sets the given name of the contact.
    /// </summary>
    public string GivenName
    {
        get => (string)PropertyBag[ContactSchema.GivenName];
        set => PropertyBag[ContactSchema.GivenName] = value;
    }

    /// <summary>
    ///     Gets or sets the initials of the contact.
    /// </summary>
    public string Initials
    {
        get => (string)PropertyBag[ContactSchema.Initials];
        set => PropertyBag[ContactSchema.Initials] = value;
    }

    /// <summary>
    ///     Gets or sets the initials of the contact.
    /// </summary>
    public string MiddleName
    {
        get => (string)PropertyBag[ContactSchema.MiddleName];
        set => PropertyBag[ContactSchema.MiddleName] = value;
    }

    /// <summary>
    ///     Gets or sets the middle name of the contact.
    /// </summary>
    public string NickName
    {
        get => (string)PropertyBag[ContactSchema.NickName];
        set => PropertyBag[ContactSchema.NickName] = value;
    }

    /// <summary>
    ///     Gets the complete name of the contact.
    /// </summary>
    public CompleteName CompleteName => (CompleteName)PropertyBag[ContactSchema.CompleteName];

    /// <summary>
    ///     Gets or sets the company name of the contact.
    /// </summary>
    public string CompanyName
    {
        get => (string)PropertyBag[ContactSchema.CompanyName];
        set => PropertyBag[ContactSchema.CompanyName] = value;
    }

    /// <summary>
    ///     Gets an indexed list of e-mail addresses for the contact. For example, to set the first e-mail address,
    ///     use the following syntax: EmailAddresses[EmailAddressKey.EmailAddress1] = "john.doe@contoso.com"
    /// </summary>
    public EmailAddressDictionary EmailAddresses => (EmailAddressDictionary)PropertyBag[ContactSchema.EmailAddresses];

    /// <summary>
    ///     Gets an indexed list of physical addresses for the contact. For example, to set the business address,
    ///     use the following syntax: PhysicalAddresses[PhysicalAddressKey.Business] = new PhysicalAddressEntry()
    /// </summary>
    public PhysicalAddressDictionary PhysicalAddresses =>
        (PhysicalAddressDictionary)PropertyBag[ContactSchema.PhysicalAddresses];

    /// <summary>
    ///     Gets an indexed list of phone numbers for the contact. For example, to set the home phone number,
    ///     use the following syntax: PhoneNumbers[PhoneNumberKey.HomePhone] = "phone number"
    /// </summary>
    public PhoneNumberDictionary PhoneNumbers => (PhoneNumberDictionary)PropertyBag[ContactSchema.PhoneNumbers];

    /// <summary>
    ///     Gets or sets the contact's assistant name.
    /// </summary>
    public string AssistantName
    {
        get => (string)PropertyBag[ContactSchema.AssistantName];
        set => PropertyBag[ContactSchema.AssistantName] = value;
    }

    /// <summary>
    ///     Gets or sets the birthday of the contact.
    /// </summary>
    public DateTime Birthday
    {
        get => (DateTime)PropertyBag[ContactSchema.Birthday];
        set => PropertyBag[ContactSchema.Birthday] = value;
    }

    /// <summary>
    ///     Gets or sets the business home page of the contact.
    /// </summary>
    public string BusinessHomePage
    {
        get => (string)PropertyBag[ContactSchema.BusinessHomePage];
        set => PropertyBag[ContactSchema.BusinessHomePage] = value;
    }

    /// <summary>
    ///     Gets or sets a list of children for the contact.
    /// </summary>
    public StringList Children
    {
        get => (StringList)PropertyBag[ContactSchema.Children];
        set => PropertyBag[ContactSchema.Children] = value;
    }

    /// <summary>
    ///     Gets or sets a list of companies for the contact.
    /// </summary>
    public StringList Companies
    {
        get => (StringList)PropertyBag[ContactSchema.Companies];
        set => PropertyBag[ContactSchema.Companies] = value;
    }

    /// <summary>
    ///     Gets the source of the contact.
    /// </summary>
    public ContactSource ContactSource => (ContactSource)PropertyBag[ContactSchema.ContactSource];

    /// <summary>
    ///     Gets or sets the department of the contact.
    /// </summary>
    public string Department
    {
        get => (string)PropertyBag[ContactSchema.Department];
        set => PropertyBag[ContactSchema.Department] = value;
    }

    /// <summary>
    ///     Gets or sets the generation of the contact.
    /// </summary>
    public string Generation
    {
        get => (string)PropertyBag[ContactSchema.Generation];
        set => PropertyBag[ContactSchema.Generation] = value;
    }

    /// <summary>
    ///     Gets an indexed list of Instant Messaging addresses for the contact. For example, to set the first
    ///     IM address, use the following syntax: ImAddresses[ImAddressKey.ImAddress1] = "john.doe@contoso.com"
    /// </summary>
    public ImAddressDictionary ImAddresses => (ImAddressDictionary)PropertyBag[ContactSchema.ImAddresses];

    /// <summary>
    ///     Gets or sets the contact's job title.
    /// </summary>
    public string JobTitle
    {
        get => (string)PropertyBag[ContactSchema.JobTitle];
        set => PropertyBag[ContactSchema.JobTitle] = value;
    }

    /// <summary>
    ///     Gets or sets the name of the contact's manager.
    /// </summary>
    public string Manager
    {
        get => (string)PropertyBag[ContactSchema.Manager];
        set => PropertyBag[ContactSchema.Manager] = value;
    }

    /// <summary>
    ///     Gets or sets the mileage for the contact.
    /// </summary>
    public string Mileage
    {
        get => (string)PropertyBag[ContactSchema.Mileage];
        set => PropertyBag[ContactSchema.Mileage] = value;
    }

    /// <summary>
    ///     Gets or sets the location of the contact's office.
    /// </summary>
    public string OfficeLocation
    {
        get => (string)PropertyBag[ContactSchema.OfficeLocation];
        set => PropertyBag[ContactSchema.OfficeLocation] = value;
    }

    /// <summary>
    ///     Gets or sets the index of the contact's postal address. When set, PostalAddressIndex refers to
    ///     an entry in the PhysicalAddresses indexed list.
    /// </summary>
    public PhysicalAddressIndex PostalAddressIndex
    {
        get => (PhysicalAddressIndex)PropertyBag[ContactSchema.PostalAddressIndex];
        set => PropertyBag[ContactSchema.PostalAddressIndex] = value;
    }

    /// <summary>
    ///     Gets or sets the contact's profession.
    /// </summary>
    public string Profession
    {
        get => (string)PropertyBag[ContactSchema.Profession];
        set => PropertyBag[ContactSchema.Profession] = value;
    }

    /// <summary>
    ///     Gets or sets the name of the contact's spouse.
    /// </summary>
    public string SpouseName
    {
        get => (string)PropertyBag[ContactSchema.SpouseName];
        set => PropertyBag[ContactSchema.SpouseName] = value;
    }

    /// <summary>
    ///     Gets or sets the surname of the contact.
    /// </summary>
    public string Surname
    {
        get => (string)PropertyBag[ContactSchema.Surname];
        set => PropertyBag[ContactSchema.Surname] = value;
    }

    /// <summary>
    ///     Gets or sets the date of the contact's wedding anniversary.
    /// </summary>
    public DateTime WeddingAnniversary
    {
        get => (DateTime)PropertyBag[ContactSchema.WeddingAnniversary];
        set => PropertyBag[ContactSchema.WeddingAnniversary] = value;
    }

    /// <summary>
    ///     Gets a value indicating whether this contact has a picture associated with it.
    /// </summary>
    public bool HasPicture => (bool)PropertyBag[ContactSchema.HasPicture];

    /// <summary>
    ///     Gets the full phonetic name from the directory
    /// </summary>
    public string PhoneticFullName => (string)PropertyBag[ContactSchema.PhoneticFullName];

    /// <summary>
    ///     Gets the phonetic first name from the directory
    /// </summary>
    public string PhoneticFirstName => (string)PropertyBag[ContactSchema.PhoneticFirstName];

    /// <summary>
    ///     Gets the phonetic last name from the directory
    /// </summary>
    public string PhoneticLastName => (string)PropertyBag[ContactSchema.PhoneticLastName];

    /// <summary>
    ///     Gets the Alias from the directory
    /// </summary>
    public string Alias => (string)PropertyBag[ContactSchema.Alias];

    /// <summary>
    ///     Get the Notes from the directory
    /// </summary>
    public string Notes => (string)PropertyBag[ContactSchema.Notes];

    /// <summary>
    ///     Gets the Photo from the directory
    /// </summary>
    public byte[] DirectoryPhoto => (byte[])PropertyBag[ContactSchema.Photo];

    /// <summary>
    ///     Gets the User SMIME certificate from the directory
    /// </summary>
    public byte[][] UserSMIMECertificate
    {
        get
        {
            var array = (ByteArrayArray)PropertyBag[ContactSchema.UserSMIMECertificate];
            return array.Content;
        }
    }

    /// <summary>
    ///     Gets the MSExchange certificate from the directory
    /// </summary>
    public byte[][] MSExchangeCertificate
    {
        get
        {
            var array = (ByteArrayArray)PropertyBag[ContactSchema.MSExchangeCertificate];
            return array.Content;
        }
    }

    /// <summary>
    ///     Gets the DirectoryID as Guid or DN string
    /// </summary>
    public string DirectoryId => (string)PropertyBag[ContactSchema.DirectoryId];

    /// <summary>
    ///     Gets the manager mailbox information
    /// </summary>
    public EmailAddress ManagerMailbox => (EmailAddress)PropertyBag[ContactSchema.ManagerMailbox];

    /// <summary>
    ///     Get the direct reports mailbox information
    /// </summary>
    public EmailAddressCollection DirectReports => (EmailAddressCollection)PropertyBag[ContactSchema.DirectReports];

    #endregion
}
