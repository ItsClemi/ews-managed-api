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
///     Represents a folder Id provided by a FolderId object.
/// </summary>
internal class FolderIdWrapper : AbstractFolderIdWrapper
{
    /// <summary>
    ///     The FolderId object providing the Id.
    /// </summary>
    private readonly FolderId _folderId;

    /// <summary>
    ///     Initializes a new instance of FolderIdWrapper.
    /// </summary>
    /// <param name="folderId">The FolderId object providing the Id.</param>
    internal FolderIdWrapper(FolderId folderId)
    {
        EwsUtilities.Assert(folderId != null, "FolderIdWrapper.ctor", "folderId is null");

        _folderId = folderId;
    }

    /// <summary>
    ///     Writes the Id encapsulated in the wrapper to XML.
    /// </summary>
    /// <param name="writer">The writer to write the Id to.</param>
    internal override void WriteToXml(EwsServiceXmlWriter writer)
    {
        _folderId.WriteToXml(writer);
    }

    /// <summary>
    ///     Validates folderId against specified version.
    /// </summary>
    /// <param name="version">The version.</param>
    internal override void Validate(ExchangeVersion version)
    {
        _folderId.Validate(version);
    }
}
