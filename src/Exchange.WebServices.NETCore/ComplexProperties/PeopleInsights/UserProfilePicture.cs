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
///     Represents the UserProfilePicture.
/// </summary>
[PublicAPI]
public sealed class UserProfilePicture : InsightValue
{
    private string _blob;
    private string _imageType;
    private string _photoSize;
    private string _url;

    /// <summary>
    ///     Gets the Blob
    /// </summary>
    public string Blob
    {
        get => _blob;
        set => SetFieldValue(ref _blob, value);
    }

    /// <summary>
    ///     Gets the PhotoSize
    /// </summary>
    public string PhotoSize
    {
        get => _photoSize;
        set => SetFieldValue(ref _photoSize, value);
    }

    /// <summary>
    ///     Gets the Url
    /// </summary>
    public string Url
    {
        get => _url;
        set => SetFieldValue(ref _url, value);
    }

    /// <summary>
    ///     Gets the ImageType
    /// </summary>
    public string ImageType
    {
        get => _imageType;
        set => SetFieldValue(ref _imageType, value);
    }

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
            case XmlElementNames.Blob:
            {
                Blob = reader.ReadElementValue();
                break;
            }
            case XmlElementNames.PhotoSize:
            {
                PhotoSize = reader.ReadElementValue();
                break;
            }
            case XmlElementNames.Url:
            {
                Url = reader.ReadElementValue();
                break;
            }
            case XmlElementNames.ImageType:
            {
                ImageType = reader.ReadElementValue();
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
