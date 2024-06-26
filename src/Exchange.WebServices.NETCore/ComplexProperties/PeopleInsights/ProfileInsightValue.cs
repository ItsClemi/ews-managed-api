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
///     Represents the ProfileInsightValue.
/// </summary>
[PublicAPI]
public sealed class ProfileInsightValue : InsightValue
{
    /// <summary>
    ///     Gets the FullName
    /// </summary>
    public string FullName { get; private set; }

    /// <summary>
    ///     Gets the FirstName
    /// </summary>
    public string FirstName { get; private set; }

    /// <summary>
    ///     Gets the LastName
    /// </summary>
    public string LastName { get; private set; }

    /// <summary>
    ///     Gets the EmailAddress
    /// </summary>
    public string EmailAddress { get; private set; }

    /// <summary>
    ///     Gets the Avatar
    /// </summary>
    public string Avatar { get; private set; }

    /// <summary>
    ///     Gets the JoinedUtcTicks
    /// </summary>
    public long JoinedUtcTicks { get; private set; }

    /// <summary>
    ///     Gets the ProfilePicture
    /// </summary>
    public UserProfilePicture ProfilePicture { get; private set; }

    /// <summary>
    ///     Gets the Title
    /// </summary>
    public string Title { get; private set; }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">XML reader</param>
    /// <returns>Whether the element was read</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.InsightSource:
            {
                InsightSource = reader.ReadElementValue<string>();
                break;
            }
            case XmlElementNames.UpdatedUtcTicks:
            {
                UpdatedUtcTicks = reader.ReadElementValue<long>();
                break;
            }
            case XmlElementNames.FullName:
            {
                FullName = reader.ReadElementValue();
                break;
            }
            case XmlElementNames.FirstName:
            {
                FirstName = reader.ReadElementValue();
                break;
            }
            case XmlElementNames.LastName:
            {
                LastName = reader.ReadElementValue();
                break;
            }
            case XmlElementNames.EmailAddress:
            {
                EmailAddress = reader.ReadElementValue();
                break;
            }
            case XmlElementNames.Avatar:
            {
                Avatar = reader.ReadElementValue();
                break;
            }
            case XmlElementNames.JoinedUtcTicks:
            {
                JoinedUtcTicks = reader.ReadElementValue<long>();
                break;
            }
            case XmlElementNames.ProfilePicture:
            {
                var picture = new UserProfilePicture();
                picture.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.ProfilePicture);
                ProfilePicture = picture;
                break;
            }
            case XmlElementNames.Title:
            {
                Title = reader.ReadElementValue();
                break;
            }
            default:
            {
                return false;
            }
        }

        return true;
    }
}
