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

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents response to UpdateFolder request.
/// </summary>
internal sealed class UpdateFolderResponse : ServiceResponse
{
    private readonly Folder _folder;

    /// <summary>
    ///     Initializes a new instance of the <see cref="UpdateFolderResponse" /> class.
    /// </summary>
    /// <param name="folder">The folder.</param>
    internal UpdateFolderResponse(Folder folder)
    {
        EwsUtilities.Assert(folder != null, "UpdateFolderResponse.ctor", "folder is null");

        _folder = folder;
    }

    /// <summary>
    ///     Reads response elements from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
    {
        base.ReadElementsFromXml(reader);

        reader.ReadServiceObjectsCollectionFromXml(XmlElementNames.Folders, GetObjectInstance, false, null, false);
    }

    /// <summary>
    ///     Clears the change log of the updated folder if the update succeeded.
    /// </summary>
    internal override void Loaded()
    {
        if (Result == ServiceResult.Success)
        {
            _folder.ClearChangeLog();
        }
    }

    /// <summary>
    ///     Gets Folder instance.
    /// </summary>
    /// <param name="session">The session.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <returns>Folder.</returns>
    private Folder GetObjectInstance(ExchangeService session, string xmlElementName)
    {
        return _folder;
    }
}
