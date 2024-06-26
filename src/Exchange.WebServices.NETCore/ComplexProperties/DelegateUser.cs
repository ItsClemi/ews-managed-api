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
///     Represents a delegate user.
/// </summary>
[PublicAPI]
public sealed class DelegateUser : ComplexProperty
{
    /// <summary>
    ///     Gets the user Id of the delegate user.
    /// </summary>
    public UserId UserId { get; private set; } = new();

    /// <summary>
    ///     Gets the list of delegate user's permissions.
    /// </summary>
    public DelegatePermissions Permissions { get; } = new();

    /// <summary>
    ///     Gets or sets a value indicating if the delegate user should receive copies of meeting requests.
    /// </summary>
    public bool ReceiveCopiesOfMeetingMessages { get; set; }

    /// <summary>
    ///     Gets or sets a value indicating if the delegate user should be able to view the principal's private items.
    /// </summary>
    public bool ViewPrivateItems { get; set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="DelegateUser" /> class.
    /// </summary>
    public DelegateUser()
    {
        // Confusing error message refers to Calendar folder permissions when adding delegate access for a user
        // without including Calendar Folder permissions.
        //
        ReceiveCopiesOfMeetingMessages = false;
        ViewPrivateItems = false;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="DelegateUser" /> class.
    /// </summary>
    /// <param name="primarySmtpAddress">The primary SMTP address of the delegate user.</param>
    public DelegateUser(string primarySmtpAddress)
        : this()
    {
        UserId.PrimarySmtpAddress = primarySmtpAddress;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="DelegateUser" /> class.
    /// </summary>
    /// <param name="standardUser">The standard delegate user.</param>
    public DelegateUser(StandardUser standardUser)
        : this()
    {
        UserId.StandardUser = standardUser;
    }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>Returns true if element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.UserId:
            {
                UserId = new UserId();
                UserId.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.DelegatePermissions:
            {
                Permissions.Reset();
                Permissions.LoadFromXml(reader, reader.LocalName);
                return true;
            }
            case XmlElementNames.ReceiveCopiesOfMeetingMessages:
            {
                ReceiveCopiesOfMeetingMessages = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.ViewPrivateItems:
            {
                ViewPrivateItems = reader.ReadElementValue<bool>();
                return true;
            }
            default:
            {
                return false;
            }
        }
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        UserId.WriteToXml(writer, XmlElementNames.UserId);
        Permissions.WriteToXml(writer, XmlElementNames.DelegatePermissions);

        writer.WriteElementValue(
            XmlNamespace.Types,
            XmlElementNames.ReceiveCopiesOfMeetingMessages,
            ReceiveCopiesOfMeetingMessages
        );

        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.ViewPrivateItems, ViewPrivateItems);
    }

    /// <summary>
    ///     Validates this instance.
    /// </summary>
    internal override void InternalValidate()
    {
        if (UserId == null)
        {
            throw new ServiceValidationException(Strings.UserIdForDelegateUserNotSpecified);
        }

        if (!UserId.IsValid())
        {
            throw new ServiceValidationException(Strings.DelegateUserHasInvalidUserId);
        }
    }

    /// <summary>
    ///     Validates this instance for AddDelegate.
    /// </summary>
    internal void ValidateAddDelegate()
    {
        Permissions.ValidateAddDelegate();
    }

    /// <summary>
    ///     Validates this instance for UpdateDelegate.
    /// </summary>
    internal void ValidateUpdateDelegate()
    {
        Permissions.ValidateUpdateDelegate();
    }
}
