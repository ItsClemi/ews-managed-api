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
///     Represents the response to an individual folder retrieval operation.
/// </summary>
[PublicAPI]
public sealed class GetFolderResponse : ServiceResponse
{
    private readonly PropertySet _propertySet;

    /// <summary>
    ///     Gets the folder that was retrieved.
    /// </summary>
    public Folder? Folder { get; private set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="GetFolderResponse" /> class.
    /// </summary>
    /// <param name="folder">The folder.</param>
    /// <param name="propertySet">The property set from the request.</param>
    internal GetFolderResponse(Folder? folder, PropertySet propertySet)
    {
        Folder = folder;
        _propertySet = propertySet;

        EwsUtilities.Assert(_propertySet != null, "GetFolderResponse.ctor", "PropertySet should not be null");
    }

    /// <summary>
    ///     Reads response elements from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
    {
        base.ReadElementsFromXml(reader);

        var folders = reader.ReadServiceObjectsCollectionFromXml(
            XmlElementNames.Folders,
            GetObjectInstance,
            true,
            _propertySet,
            false
        );

        Folder = folders[0];
    }

    /// <summary>
    ///     Gets the folder instance.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <returns>Folder.</returns>
    private Folder? GetObjectInstance(ExchangeService service, string xmlElementName)
    {
        if (Folder != null)
        {
            return Folder;
        }

        return EwsUtilities.CreateEwsObjectFromXmlElementName<Folder>(service, xmlElementName);
    }
}
