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

using System.Xml;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a app in GetAppManifests response.
/// </summary>
[PublicAPI]
public sealed class ClientApp : ComplexProperty
{
    /// <summary>
    ///     The manifest for the app.
    /// </summary>
    public XmlDocument Manifest { get; internal set; }

    /// <summary>
    ///     Metadata related to the app.
    /// </summary>
    public ClientAppMetadata Metadata { get; internal set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ClientApp" /> class.
    /// </summary>
    internal ClientApp()
    {
        Namespace = XmlNamespace.Types;
    }

    /// <summary>
    ///     Helper to convert to xml document from the current value.
    /// </summary>
    /// <param name="reader">the reader.</param>
    /// <returns>The xml document</returns>
    internal static SafeXmlDocument ReadToXmlDocument(EwsServiceXmlReader reader)
    {
        using var stream = new MemoryStream();
        reader.ReadBase64ElementValue(stream);
        stream.Position = 0;

        var manifest = new SafeXmlDocument();
        manifest.Load(stream);
        return manifest;
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
            case XmlElementNames.Manifest:
            {
                Manifest = ReadToXmlDocument(reader);
                return true;
            }
            case XmlElementNames.Metadata:
            {
                Metadata = new ClientAppMetadata();
                Metadata.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.Metadata);
                return true;
            }
            default:
            {
                return false;
            }
        }
    }
}
