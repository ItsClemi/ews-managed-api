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
using System.Xml.Schema;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     XmlSchema with protection against DTD parsing in read overloads.
/// </summary>
internal class SafeXmlSchema : XmlSchema
{
    /// <summary>
    ///     Safe xml reader settings.
    /// </summary>
    private static readonly XmlReaderSettings DefaultSettings = new()
    {
        DtdProcessing = DtdProcessing.Prohibit,
    };


    #region Methods

    /// <summary>
    ///     Reads an XML Schema from the supplied stream.
    /// </summary>
    /// <param name="stream">The supplied data stream.</param>
    /// <param name="validationEventHandler">
    ///     The validation event handler that receives information about the XML Schema syntax
    ///     errors.
    /// </param>
    /// <returns>The XmlSchema object representing the XML Schema.</returns>
    public new static XmlSchema? Read(Stream stream, ValidationEventHandler validationEventHandler)
    {
        using var xr = XmlReader.Create(stream, DefaultSettings);
        return XmlSchema.Read(xr, validationEventHandler);
    }

    /// <summary>
    ///     Reads an XML Schema from the supplied TextReader.
    /// </summary>
    /// <param name="reader">The TextReader containing the XML Schema to read.</param>
    /// <param name="validationEventHandler">
    ///     The validation event handler that receives information about the XML Schema syntax
    ///     errors.
    /// </param>
    /// <returns>The XmlSchema object representing the XML Schema.</returns>
    public new static XmlSchema? Read(TextReader reader, ValidationEventHandler validationEventHandler)
    {
        using var xr = XmlReader.Create(reader, DefaultSettings);
        return XmlSchema.Read(xr, validationEventHandler);
    }

    /// <summary>
    ///     Reads an XML Schema from the supplied XmlReader.
    /// </summary>
    /// <param name="reader">The XmlReader containing the XML Schema to read.</param>
    /// <param name="validationEventHandler">
    ///     The validation event handler that receives information about the XML Schema syntax
    ///     errors.
    /// </param>
    /// <returns>The XmlSchema object representing the XML Schema.</returns>
    public new static XmlSchema? Read(XmlReader reader, ValidationEventHandler validationEventHandler)
    {
        // we need to check to see if the reader is configured properly
        if (reader.Settings != null)
        {
            if (reader.Settings.DtdProcessing != DtdProcessing.Prohibit)
            {
                throw new XmlDtdException();
            }
        }

        return XmlSchema.Read(reader, validationEventHandler);
    }

    #endregion
}
