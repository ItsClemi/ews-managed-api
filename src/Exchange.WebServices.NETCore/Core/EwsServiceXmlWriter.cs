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

using System.Globalization;
using System.Reflection;
using System.Text;
using System.Xml;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     XML writer
/// </summary>
internal class EwsServiceXmlWriter : IDisposable
{
    /// <summary>
    ///     Buffer size for writing Base64 encoded content.
    /// </summary>
    private const int BufferSize = 8192;

    /// <summary>
    ///     UTF-8 encoding that does not create leading Byte order marks
    /// </summary>
    private static readonly Encoding Utf8Encoding = new UTF8Encoding(false);

    private bool _isDisposed;

    /// <summary>
    ///     Gets the internal XML writer.
    /// </summary>
    /// <value>The internal writer.</value>
    public XmlWriter InternalWriter { get; }

    /// <summary>
    ///     Gets the service.
    /// </summary>
    /// <value>The service.</value>
    public ExchangeServiceBase Service { get; }

    /// <summary>
    ///     Gets or sets a value indicating whether the time zone SOAP header was emitted through this writer.
    /// </summary>
    /// <value>
    ///     <c>true</c> if the time zone SOAP header was emitted; otherwise, <c>false</c>.
    /// </value>
    public bool IsTimeZoneHeaderEmitted { get; set; }

    /// <summary>
    ///     Gets or sets a value indicating whether the SOAP message need WSSecurity Utility namespace.
    /// </summary>
    public bool RequireWsSecurityUtilityNamespace { get; set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="EwsServiceXmlWriter" /> class.
    /// </summary>
    /// <param name="service">The service.</param>
    /// <param name="stream">The stream.</param>
    internal EwsServiceXmlWriter(ExchangeServiceBase service, Stream stream)
    {
        Service = service;

        var settings = new XmlWriterSettings
        {
            Indent = true,
            Encoding = Utf8Encoding,
        };

        InternalWriter = XmlWriter.Create(stream, settings);
    }

    /// <summary>
    ///     Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
    /// </summary>
    public void Dispose()
    {
        if (!_isDisposed)
        {
            InternalWriter.Dispose();

            _isDisposed = true;
        }
    }

    /// <summary>
    ///     Try to convert object to a string.
    /// </summary>
    /// <param name="value">The value.</param>
    /// <param name="strValue">The string representation of value.</param>
    /// <returns>True if object was converted, false otherwise.</returns>
    /// <remarks>A null object will be "successfully" converted to a null string.</remarks>
    internal bool TryConvertObjectToString(object? value, out string? strValue)
    {
        strValue = null;

        if (value == null)
        {
            return true;
        }

        // All value types should implement IConvertible. There are a couple of special cases 
        // that need to be handled directly. Otherwise use IConvertible.ToString()
        if (value.GetType().GetTypeInfo().IsEnum)
        {
            strValue = EwsUtilities.SerializeEnum((Enum)value);
            return true;
        }

        if (value is IConvertible convertible)
        {
            strValue = convertible.GetTypeCode() switch
            {
                TypeCode.Boolean => EwsUtilities.BoolToXsBool((bool)value),
                TypeCode.DateTime => Service.ConvertDateTimeToUniversalDateTimeString((DateTime)value),
                _ => convertible.ToString(CultureInfo.InvariantCulture),
            };

            return true;
        }

        switch (value)
        {
            // If the value type doesn't implement IConvertible but implements IFormattable, use its
            // ToString(format,formatProvider) method to convert to a string.
            case IFormattable formattable:
            {
                // Null arguments mean that we use default format and default locale.
                strValue = formattable.ToString(null, null);
                break;
            }
            case ISearchStringProvider searchStringProvider:
            {
                // If the value type doesn't implement IConvertible or IFormattable but implements 
                // ISearchStringProvider convert to a string.
                // Note: if a value type implements IConvertible or IFormattable we will *not* check
                // to see if it also implements ISearchStringProvider. We'll always use its IConvertible.ToString 
                // or IFormattable.ToString method.
                strValue = searchStringProvider.GetSearchString();
                break;
            }
            case byte[] bytes:
            {
                // Special case for byte arrays. Convert to Base64-encoded string.
                strValue = Convert.ToBase64String(bytes);
                break;
            }
            default:
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    ///     Flushes this instance.
    /// </summary>
    public void Flush()
    {
        InternalWriter.Flush();
    }

    /// <summary>
    ///     Writes the start element.
    /// </summary>
    /// <param name="xmlNamespace">The XML namespace.</param>
    /// <param name="localName">The local name of the element.</param>
    public void WriteStartElement(XmlNamespace xmlNamespace, string localName)
    {
        InternalWriter.WriteStartElement(
            EwsUtilities.GetNamespacePrefix(xmlNamespace),
            localName,
            EwsUtilities.GetNamespaceUri(xmlNamespace)
        );
    }

    /// <summary>
    ///     Writes the end element.
    /// </summary>
    public void WriteEndElement()
    {
        InternalWriter.WriteEndElement();
    }

    /// <summary>
    ///     Writes the attribute value.  Does not emit empty string values.
    /// </summary>
    /// <param name="localName">The local name of the attribute.</param>
    /// <param name="value">The value.</param>
    public void WriteAttributeValue(string localName, object value)
    {
        WriteAttributeValue(localName, false, value);
    }

    /// <summary>
    ///     Writes the attribute value.  Optionally emits empty string values.
    /// </summary>
    /// <param name="localName">The local name of the attribute.</param>
    /// <param name="alwaysWriteEmptyString">Always emit the empty string as the value.</param>
    /// <param name="value">The value.</param>
    public void WriteAttributeValue(string localName, bool alwaysWriteEmptyString, object? value)
    {
        if (!TryConvertObjectToString(value, out var stringValue))
        {
            throw new ServiceXmlSerializationException(
                string.Format(Strings.AttributeValueCannotBeSerialized, value!.GetType().Name, localName)
            );
        }

        if (stringValue != null && (alwaysWriteEmptyString || stringValue.Length != 0))
        {
            WriteAttributeString(localName, stringValue);
        }
    }

    /// <summary>
    ///     Writes the attribute value.
    /// </summary>
    /// <param name="namespacePrefix">The namespace prefix.</param>
    /// <param name="localName">The local name of the attribute.</param>
    /// <param name="value">The value.</param>
    public void WriteAttributeValue(string namespacePrefix, string localName, object value)
    {
        if (!TryConvertObjectToString(value, out var stringValue))
        {
            throw new ServiceXmlSerializationException(
                string.Format(Strings.AttributeValueCannotBeSerialized, value.GetType().Name, localName)
            );
        }

        if (!string.IsNullOrEmpty(stringValue))
        {
            WriteAttributeString(namespacePrefix, localName, stringValue);
        }
    }

    /// <summary>
    ///     Writes the attribute value.
    /// </summary>
    /// <param name="localName">The local name of the attribute.</param>
    /// <param name="stringValue">The string value.</param>
    /// <exception cref="ServiceXmlSerializationException">Thrown if string value isn't valid for XML.</exception>
    internal void WriteAttributeString(string localName, string stringValue)
    {
        try
        {
            InternalWriter.WriteAttributeString(localName, stringValue);
        }
        catch (ArgumentException ex)
        {
            // XmlTextWriter will throw ArgumentException if string includes invalid characters.
            throw new ServiceXmlSerializationException(
                string.Format(Strings.InvalidAttributeValue, stringValue, localName),
                ex
            );
        }
    }

    /// <summary>
    ///     Writes the attribute value.
    /// </summary>
    /// <param name="namespacePrefix">The namespace prefix.</param>
    /// <param name="localName">The local name of the attribute.</param>
    /// <param name="stringValue">The string value.</param>
    /// <exception cref="ServiceXmlSerializationException">Thrown if string value isn't valid for XML.</exception>
    internal void WriteAttributeString(string namespacePrefix, string localName, string stringValue)
    {
        try
        {
            InternalWriter.WriteAttributeString(namespacePrefix, localName, null, stringValue);
        }
        catch (ArgumentException ex)
        {
            // XmlTextWriter will throw ArgumentException if string includes invalid characters.
            throw new ServiceXmlSerializationException(
                string.Format(Strings.InvalidAttributeValue, stringValue, localName),
                ex
            );
        }
    }

    /// <summary>
    ///     Writes string value.
    /// </summary>
    /// <param name="value">The value.</param>
    /// <param name="name">Element name (used for error handling)</param>
    /// <exception cref="ServiceXmlSerializationException">Thrown if string value isn't valid for XML.</exception>
    public void WriteValue(string value, string name)
    {
        try
        {
            InternalWriter.WriteValue(value);
        }
        catch (ArgumentException ex)
        {
            // XmlTextWriter will throw ArgumentException if string includes invalid characters.
            throw new ServiceXmlSerializationException(
                string.Format(Strings.InvalidElementStringValue, value, name),
                ex
            );
        }
    }

    /// <summary>
    ///     Writes the element value.
    /// </summary>
    /// <param name="xmlNamespace">The XML namespace.</param>
    /// <param name="localName">The local name of the element.</param>
    /// <param name="displayName">The name that should appear in the exception message when the value can not be serialized.</param>
    /// <param name="value">The value.</param>
    internal void WriteElementValue(XmlNamespace xmlNamespace, string localName, string displayName, object? value)
    {
        if (!TryConvertObjectToString(value, out var stringValue))
        {
            throw new ServiceXmlSerializationException(
                string.Format(Strings.ElementValueCannotBeSerialized, value.GetType().Name, localName)
            );
        }

        //  PS # 205106: The code here used to check IsNullOrEmpty on stringValue instead of just null.
        //  Unfortunately, that meant that if someone really needed to update a string property to be the
        //  value "" (String.Empty), they couldn't do it, because we wouldn't emit the element here, causing
        //  an error on the server because an update is required to have a single sub-element that is the
        //  value to update.  So we need to allow an empty string to create an empty element (like <Value />).
        //  Note that changing this check to just check for null is fine, because the other types that get
        //  converted by TryConvertObjectToString() won't return an empty string if the conversion is
        //  successful (for instance, converting an integer to a string won't return an empty string - it'll
        //  always return the stringized integer).
        if (stringValue != null)
        {
            WriteStartElement(xmlNamespace, localName);
            WriteValue(stringValue, displayName);
            WriteEndElement();
        }
    }

    /// <summary>
    ///     Writes the Xml Node
    /// </summary>
    /// <param name="xmlNode">The XML node.</param>
    public void WriteNode(XmlNode? xmlNode)
    {
        xmlNode?.WriteTo(InternalWriter);
    }

    /// <summary>
    ///     Writes the element value.
    /// </summary>
    /// <param name="xmlNamespace">The XML namespace.</param>
    /// <param name="localName">The local name of the element.</param>
    /// <param name="value">The value.</param>
    public void WriteElementValue(XmlNamespace xmlNamespace, string localName, object? value)
    {
        WriteElementValue(xmlNamespace, localName, localName, value);
    }

    /// <summary>
    ///     Writes the base64-encoded element value.
    /// </summary>
    /// <param name="buffer">The buffer.</param>
    public void WriteBase64ElementValue(byte[] buffer)
    {
        InternalWriter.WriteBase64(buffer, 0, buffer.Length);
    }

    /// <summary>
    ///     Writes the base64-encoded element value.
    /// </summary>
    /// <param name="stream">The stream.</param>
    public void WriteBase64ElementValue(Stream stream)
    {
        var buffer = new byte[BufferSize];

        using var reader = new BinaryReader(stream);
        int bytesRead;
        do
        {
            bytesRead = reader.Read(buffer, 0, BufferSize);

            if (bytesRead > 0)
            {
                InternalWriter.WriteBase64(buffer, 0, bytesRead);
            }
        } while (bytesRead > 0);
    }
}
