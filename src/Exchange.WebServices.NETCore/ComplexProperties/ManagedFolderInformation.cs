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
///     Represents information for a managed folder.
/// </summary>
[PublicAPI]
public sealed class ManagedFolderInformation : ComplexProperty
{
    private string _comment;
    private string _homePage;

    /// <summary>
    ///     Gets a value indicating whether the user can delete objects in the folder.
    /// </summary>
    public bool? CanDelete { get; private set; }

    /// <summary>
    ///     Gets a value indicating whether the user can rename or move objects in the folder.
    /// </summary>
    public bool? CanRenameOrMove { get; private set; }

    /// <summary>
    ///     Gets a value indicating whether the client application must display the Comment property to the user.
    /// </summary>
    public bool? MustDisplayComment { get; private set; }

    /// <summary>
    ///     Gets a value indicating whether the folder has a quota.
    /// </summary>
    public bool? HasQuota { get; private set; }

    /// <summary>
    ///     Gets a value indicating whether the folder is the root of the managed folder hierarchy.
    /// </summary>
    public bool? IsManagedFoldersRoot { get; private set; }

    /// <summary>
    ///     Gets the Managed Folder Id of the folder.
    /// </summary>
    public string ManagedFolderId { get; private set; }

    /// <summary>
    ///     Gets the comment associated with the folder.
    /// </summary>
    public string Comment => _comment;

    /// <summary>
    ///     Gets the storage quota of the folder.
    /// </summary>
    public int? StorageQuota { get; private set; }

    /// <summary>
    ///     Gets the size of the folder.
    /// </summary>
    public int? FolderSize { get; private set; }

    /// <summary>
    ///     Gets the home page associated with the folder.
    /// </summary>
    public string HomePage => _homePage;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ManagedFolderInformation" /> class.
    /// </summary>
    internal ManagedFolderInformation()
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
            case XmlElementNames.CanDelete:
            {
                CanDelete = reader.ReadValue<bool>();
                return true;
            }
            case XmlElementNames.CanRenameOrMove:
            {
                CanRenameOrMove = reader.ReadValue<bool>();
                return true;
            }
            case XmlElementNames.MustDisplayComment:
            {
                MustDisplayComment = reader.ReadValue<bool>();
                return true;
            }
            case XmlElementNames.HasQuota:
            {
                HasQuota = reader.ReadValue<bool>();
                return true;
            }
            case XmlElementNames.IsManagedFoldersRoot:
            {
                IsManagedFoldersRoot = reader.ReadValue<bool>();
                return true;
            }
            case XmlElementNames.ManagedFolderId:
            {
                ManagedFolderId = reader.ReadValue();
                return true;
            }
            case XmlElementNames.Comment:
            {
                reader.TryReadValue(ref _comment);
                return true;
            }
            case XmlElementNames.StorageQuota:
            {
                StorageQuota = reader.ReadValue<int>();
                return true;
            }
            case XmlElementNames.FolderSize:
            {
                FolderSize = reader.ReadValue<int>();
                return true;
            }
            case XmlElementNames.HomePage:
            {
                reader.TryReadValue(ref _homePage);
                return true;
            }
            default:
            {
                return false;
            }
        }
    }
}
