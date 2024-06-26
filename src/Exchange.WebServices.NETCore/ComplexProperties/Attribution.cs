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
///     Represents an attribution of an attributed string
/// </summary>
[PublicAPI]
public sealed class Attribution : ComplexProperty
{
    /// <summary>
    ///     Attribution id
    /// </summary>
    public string Id { get; set; }

    /// <summary>
    ///     Attribution source
    /// </summary>
    public ItemId SourceId { get; set; }

    /// <summary>
    ///     Display name
    /// </summary>
    public string DisplayName { get; set; }

    /// <summary>
    ///     Whether writable
    /// </summary>
    public bool IsWritable { get; set; }

    /// <summary>
    ///     Whether a quick contact
    /// </summary>
    public bool IsQuickContact { get; set; }

    /// <summary>
    ///     Whether hidden
    /// </summary>
    public bool IsHidden { get; set; }

    /// <summary>
    ///     Folder id
    /// </summary>
    public FolderId FolderId { get; set; }

    /// <summary>
    ///     Default constructor
    /// </summary>
    public Attribution()
    {
    }

    /// <summary>
    ///     Creates an instance with all values
    /// </summary>
    /// <param name="id">Attribution id</param>
    /// <param name="sourceId">Source Id</param>
    /// <param name="displayName">Display name</param>
    /// <param name="isWritable">Whether writable</param>
    /// <param name="isQuickContact">Wther quick contact</param>
    /// <param name="isHidden">Whether hidden</param>
    /// <param name="folderId">Folder id</param>
    public Attribution(
        string id,
        ItemId sourceId,
        string displayName,
        bool isWritable = false,
        bool isQuickContact = false,
        bool isHidden = false,
        FolderId? folderId = null
    )
        : this()
    {
        EwsUtilities.ValidateParam(id);
        EwsUtilities.ValidateParam(displayName);

        Id = id;
        SourceId = sourceId;
        DisplayName = displayName;
        IsWritable = isWritable;
        IsQuickContact = isQuickContact;
        IsHidden = isHidden;
        FolderId = folderId;
    }

    /// <summary>
    ///     Tries to read element from XML
    /// </summary>
    /// <param name="reader">XML reader</param>
    /// <returns>Whether reading succeeded</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.Id:
            {
                Id = reader.ReadElementValue();
                break;
            }
            case XmlElementNames.SourceId:
            {
                SourceId = new ItemId();
                SourceId.LoadFromXml(reader, reader.LocalName);
                break;
            }
            case XmlElementNames.DisplayName:
            {
                DisplayName = reader.ReadElementValue();
                break;
            }
            case XmlElementNames.IsWritable:
            {
                IsWritable = reader.ReadElementValue<bool>();
                break;
            }
            case XmlElementNames.IsQuickContact:
            {
                IsQuickContact = reader.ReadElementValue<bool>();
                break;
            }
            case XmlElementNames.IsHidden:
            {
                IsHidden = reader.ReadElementValue<bool>();
                break;
            }
            case XmlElementNames.FolderId:
            {
                FolderId = new FolderId();
                FolderId.LoadFromXml(reader, reader.LocalName);
                break;
            }
            default:
            {
                return base.TryReadElementFromXml(reader);
            }
        }

        return true;
    }
}
