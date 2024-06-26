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
///     Represents the Id of a user.
/// </summary>
[PublicAPI]
public sealed class UserId : ComplexProperty
{
    private string _displayName;
    private string _primarySmtpAddress;
    private string _sId;
    private StandardUser? _standardUser;

    /// <summary>
    ///     Gets or sets the SID of the user.
    /// </summary>
    public string SID
    {
        get => _sId;
        set => SetFieldValue(ref _sId, value);
    }

    /// <summary>
    ///     Gets or sets the primary SMTP address or the user.
    /// </summary>
    public string PrimarySmtpAddress
    {
        get => _primarySmtpAddress;
        set => SetFieldValue(ref _primarySmtpAddress, value);
    }

    /// <summary>
    ///     Gets or sets the display name of the user.
    /// </summary>
    public string DisplayName
    {
        get => _displayName;
        set => SetFieldValue(ref _displayName, value);
    }

    /// <summary>
    ///     Gets or sets a value indicating which standard user the user represents.
    /// </summary>
    public StandardUser? StandardUser
    {
        get => _standardUser;
        set => SetFieldValue(ref _standardUser, value);
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="UserId" /> class.
    /// </summary>
    public UserId()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="UserId" /> class.
    /// </summary>
    /// <param name="primarySmtpAddress">The primary SMTP address used to initialize the UserId.</param>
    public UserId(string primarySmtpAddress)
        : this()
    {
        _primarySmtpAddress = primarySmtpAddress;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="UserId" /> class.
    /// </summary>
    /// <param name="standardUser">The StandardUser value used to initialize the UserId.</param>
    public UserId(StandardUser standardUser)
        : this()
    {
        _standardUser = standardUser;
    }

    /// <summary>
    ///     Determines whether this instance is valid.
    /// </summary>
    /// <returns><c>true</c> if this instance is valid; otherwise, <c>false</c>.</returns>
    internal bool IsValid()
    {
        return StandardUser.HasValue || !string.IsNullOrEmpty(PrimarySmtpAddress) || !string.IsNullOrEmpty(SID);
    }

    /// <summary>
    ///     Implements an implicit conversion between a string representing a primary SMTP address and UserId.
    /// </summary>
    /// <param name="primarySmtpAddress">The string representing a primary SMTP address.</param>
    /// <returns>A UserId initialized with the specified primary SMTP address.</returns>
    public static implicit operator UserId(string primarySmtpAddress)
    {
        return new UserId(primarySmtpAddress);
    }

    /// <summary>
    ///     Implements an implicit conversion between StandardUser and UserId.
    /// </summary>
    /// <param name="standardUser">The standard user used to initialize the user Id.</param>
    /// <returns>A UserId initialized with the specified standard user value.</returns>
    public static implicit operator UserId(StandardUser standardUser)
    {
        return new UserId(standardUser);
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
            case XmlElementNames.SID:
            {
                _sId = reader.ReadValue();
                return true;
            }
            case XmlElementNames.PrimarySmtpAddress:
            {
                _primarySmtpAddress = reader.ReadValue();
                return true;
            }
            case XmlElementNames.DisplayName:
            {
                _displayName = reader.ReadValue();
                return true;
            }
            case XmlElementNames.DistinguishedUser:
            {
                _standardUser = reader.ReadValue<StandardUser>();
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
        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.SID, SID);

        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.PrimarySmtpAddress, PrimarySmtpAddress);

        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.DisplayName, DisplayName);

        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.DistinguishedUser, StandardUser);
    }
}
