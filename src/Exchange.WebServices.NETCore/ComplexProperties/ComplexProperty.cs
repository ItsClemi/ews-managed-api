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

using System.ComponentModel;
using System.Xml;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a property that can be sent to or retrieved from EWS.
/// </summary>
[PublicAPI]
[EditorBrowsable(EditorBrowsableState.Never)]
public abstract class ComplexProperty : ISelfValidate
{
    /// <summary>
    ///     Gets or sets the namespace.
    /// </summary>
    /// <value>The namespace.</value>
    internal XmlNamespace Namespace { get; set; } = XmlNamespace.Types;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ComplexProperty" /> class.
    /// </summary>
    internal ComplexProperty()
    {
    }

    /// <summary>
    ///     Implements ISelfValidate.Validate. Validates this instance.
    /// </summary>
    void ISelfValidate.Validate()
    {
        InternalValidate();
    }

    /// <summary>
    ///     Instance was changed.
    /// </summary>
    internal virtual void Changed()
    {
        OnChange?.Invoke(this);
    }

    /// <summary>
    ///     Sets value of field.
    /// </summary>
    /// <typeparam name="T">Field type.</typeparam>
    /// <param name="field">The field.</param>
    /// <param name="value">The value.</param>
    internal virtual void SetFieldValue<T>(ref T field, T value)
    {
        bool applyChange;

        if (field == null)
        {
            applyChange = value != null;
        }
        else
        {
            if (field is IComparable comparable)
            {
                applyChange = comparable.CompareTo(value) != 0;
            }
            else
            {
                applyChange = true;
            }
        }

        if (applyChange)
        {
            field = value;
            Changed();
        }
    }

    /// <summary>
    ///     Clears the change log.
    /// </summary>
    internal virtual void ClearChangeLog()
    {
    }

    /// <summary>
    ///     Reads the attributes from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal virtual void ReadAttributesFromXml(EwsServiceXmlReader reader)
    {
    }

    /// <summary>
    ///     Reads the text value from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal virtual void ReadTextValueFromXml(EwsServiceXmlReader reader)
    {
    }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if element was read.</returns>
    internal virtual bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        return false;
    }

    /// <summary>
    ///     Tries to read element from XML to patch this property.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if element was read.</returns>
    internal virtual bool TryReadElementFromXmlToPatch(EwsServiceXmlReader reader)
    {
        return false;
    }

    /// <summary>
    ///     Writes the attributes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal virtual void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal virtual void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
    }

    /// <summary>
    ///     Loads from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="xmlNamespace">The XML namespace.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    internal virtual void LoadFromXml(EwsServiceXmlReader reader, XmlNamespace xmlNamespace, string xmlElementName)
    {
        InternalLoadFromXml(reader, xmlNamespace, xmlElementName, TryReadElementFromXml);
    }

    /// <summary>
    ///     Loads from XML to update itself.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="xmlNamespace">The XML namespace.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    internal virtual void UpdateFromXml(EwsServiceXmlReader reader, XmlNamespace xmlNamespace, string xmlElementName)
    {
        InternalLoadFromXml(reader, xmlNamespace, xmlElementName, TryReadElementFromXmlToPatch);
    }

    /// <summary>
    ///     Loads from XML
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="xmlNamespace">The XML namespace.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <param name="readAction"></param>
    private void InternalLoadFromXml(
        EwsServiceXmlReader reader,
        XmlNamespace xmlNamespace,
        string xmlElementName,
        Func<EwsServiceXmlReader, bool> readAction
    )
    {
        reader.EnsureCurrentNodeIsStartElement(xmlNamespace, xmlElementName);

        ReadAttributesFromXml(reader);

        if (!reader.IsEmptyElement)
        {
            do
            {
                reader.Read();

                switch (reader.NodeType)
                {
                    case XmlNodeType.Element:
                    {
                        if (!readAction(reader))
                        {
                            reader.SkipCurrentElement();
                        }

                        break;
                    }
                    case XmlNodeType.Text:
                    {
                        ReadTextValueFromXml(reader);
                        break;
                    }
                }
            } while (!reader.IsEndElement(xmlNamespace, xmlElementName));
        }
    }

    /// <summary>
    ///     Loads from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    internal virtual void LoadFromXml(EwsServiceXmlReader reader, string xmlElementName)
    {
        LoadFromXml(reader, Namespace, xmlElementName);
    }

    /// <summary>
    ///     Loads from XML to update this property.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    internal virtual void UpdateFromXml(EwsServiceXmlReader reader, string xmlElementName)
    {
        UpdateFromXml(reader, Namespace, xmlElementName);
    }

    /// <summary>
    ///     Writes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="xmlNamespace">The XML namespace.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    internal virtual void WriteToXml(EwsServiceXmlWriter writer, XmlNamespace xmlNamespace, string xmlElementName)
    {
        writer.WriteStartElement(xmlNamespace, xmlElementName);
        WriteAttributesToXml(writer);
        WriteElementsToXml(writer);
        writer.WriteEndElement();
    }

    /// <summary>
    ///     Writes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    internal virtual void WriteToXml(EwsServiceXmlWriter writer, string xmlElementName)
    {
        WriteToXml(writer, Namespace, xmlElementName);
    }

    /// <summary>
    ///     Occurs when property changed.
    /// </summary>
    internal event ComplexPropertyChangedDelegate? OnChange;

    /// <summary>
    ///     Validates this instance.
    /// </summary>
    internal virtual void InternalValidate()
    {
    }
}
