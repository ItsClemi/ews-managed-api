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
///     Represents the SetClientExtension method action.
/// </summary>
[PublicAPI]
public sealed class SetClientExtensionAction : ComplexProperty
{
    private readonly ClientExtension? _clientExtension;
    private readonly string _extensionId;
    private readonly SetClientExtensionActionId _setClientExtensionActionId;

    /// <summary>
    ///     Initializes a new instance of the <see cref="SetClientExtensionAction" /> class.
    /// </summary>
    /// <param name="setClientExtensionActionId">Set action such as install, uninstall and configure</param>
    /// <param name="extensionId">ExtensionId, required by configure and uninstall actions</param>
    /// <param name="clientExtension">Extension data object, e.g. required by configure action</param>
    public SetClientExtensionAction(
        SetClientExtensionActionId setClientExtensionActionId,
        string extensionId,
        ClientExtension clientExtension
    )
    {
        Namespace = XmlNamespace.Types;
        _setClientExtensionActionId = setClientExtensionActionId;
        _extensionId = extensionId;
        _clientExtension = clientExtension;
    }

    /// <summary>
    ///     Writes attributes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteAttributeValue(XmlAttributeNames.SetClientExtensionActionId, _setClientExtensionActionId);

        if (!string.IsNullOrEmpty(_extensionId))
        {
            writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionId, _extensionId);
        }
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        _clientExtension?.WriteToXml(writer, XmlNamespace.Types, XmlElementNames.ClientExtension);
    }
}
