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

using System.Globalization;
using System.Net.Http.Headers;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents the response to GetUserPhoto operation.
/// </summary>
internal sealed class GetUserPhotoResponse : ServiceResponse
{
    /// <summary>
    ///     Gets GetUserPhoto results.
    /// </summary>
    /// <returns>GetUserPhoto results.</returns>
    internal GetUserPhotoResults Results { get; private set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="GetUserPhotoResponse" /> class.
    /// </summary>
    internal GetUserPhotoResponse()
    {
        Results = new GetUserPhotoResults();
    }

    /// <summary>
    ///     Read Photo results from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
    {
        var hasChanged = reader.ReadElementValue<bool>(XmlNamespace.Messages, XmlElementNames.HasChanged);

        reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.PictureData);
        var photoData = reader.ReadBase64ElementValue();

        // We only get a content type if we get a photo
        if (photoData.Length > 0)
        {
            Results.Photo = photoData;
            Results.ContentType = reader.ReadElementValue(XmlNamespace.Messages, XmlElementNames.ContentType);
        }

        if (hasChanged)
        {
            if (Results.Photo.Length == 0)
            {
                Results.Status = GetUserPhotoStatus.PhotoOrUserNotFound;
            }
            else
            {
                Results.Status = GetUserPhotoStatus.PhotoReturned;
            }
        }
        else
        {
            Results.Status = GetUserPhotoStatus.PhotoUnchanged;
        }
    }

    /// <summary>
    ///     Read Photo response headers
    /// </summary>
    /// <param name="responseHeaders">The response header.</param>
    internal override void ReadHeader(HttpResponseHeaders responseHeaders)
    {
        // Parse out the ETag, trimming the quotes
        if (responseHeaders.TryGetValues("ETag", out var etagValues))
        {
            var etag = etagValues.First();
            etag = etag.Replace("\"", "");

            if (etag.Length > 0)
            {
                Results.EntityTag = etag;
            }
        }

        // Parse the Expires tag, leaving it in UTC
        if (responseHeaders.TryGetValues("Expires", out var expiresValues))
        {
            var expires = expiresValues.First();
            if (expires != null && expires.Length > 0)
            {
                Results.Expires = DateTime.Parse(expires, null, DateTimeStyles.RoundtripKind);
            }
        }
    }
}
