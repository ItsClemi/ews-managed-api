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

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     XmlDocument that does not allow DTD parsing.
/// </summary>
internal class SafeXmlDocument : XmlDocument
{
    /// <summary>
    ///     Xml settings object.
    /// </summary>
    private readonly XmlReaderSettings _settings = new()
    {
        DtdProcessing = DtdProcessing.Prohibit,
    };


    /// <summary>
    ///     Initializes a new instance of the SafeXmlDocument class.
    /// </summary>
    public SafeXmlDocument()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the SafeXmlDocument class with the specified XmlImplementation.
    /// </summary>
    /// <remarks>Not supported do to no use within exchange dev code.</remarks>
    /// <param name="imp">The XmlImplementation to use.</param>
    public SafeXmlDocument(XmlImplementation imp)
    {
        throw new NotSupportedException("Not supported");
    }

    /// <summary>
    ///     Initializes a new instance of the SafeXmlDocument class with the specified XmlNameTable.
    /// </summary>
    /// <param name="nt">The XmlNameTable to use.</param>
    public SafeXmlDocument(XmlNameTable nt)
        : base(nt)
    {
    }


    #region Methods

    /// <summary>
    ///     Loads the XML document from the specified stream.
    /// </summary>
    /// <param name="inStream">The stream containing the XML document to load.</param>
    public override void Load(Stream inStream)
    {
        using var reader = XmlReader.Create(inStream, _settings);
        Load(reader);
    }

    /// <summary>
    ///     Loads the XML document from the specified TextReader.
    /// </summary>
    /// <param name="txtReader">The TextReader used to feed the XML data into the document.</param>
    public override void Load(TextReader txtReader)
    {
        using var reader = XmlReader.Create(txtReader, _settings);
        Load(reader);
    }

    /// <summary>
    ///     Loads the XML document from the specified XmlReader.
    /// </summary>
    /// <param name="reader">The XmlReader used to feed the XML data into the document.</param>
    public override void Load(XmlReader reader)
    {
        // we need to check to see if the reader is configured properly
        if (reader.Settings != null)
        {
            if (reader.Settings.DtdProcessing != DtdProcessing.Prohibit)
            {
                throw new XmlDtdException();
            }
        }

        try
        {
            base.Load(reader);
        }
        catch (XmlException x)
        {
            if (x.Message.StartsWith(
                    "For security reasons DTD is prohibited in this XML document.",
                    StringComparison.OrdinalIgnoreCase
                ))
            {
                throw new XmlDtdException();
            }
        }
    }

    /// <summary>
    ///     Loads the XML document from the specified string.
    /// </summary>
    /// <param name="xml">String containing the XML document to load.</param>
    public override void LoadXml(string xml)
    {
        using var reader = XmlReader.Create(new StringReader(xml), _settings);
        base.Load(reader);
    }

    #endregion
}
