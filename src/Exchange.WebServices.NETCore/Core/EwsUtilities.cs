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

using System.Collections;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Net.Http.Headers;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     EWS utilities
/// </summary>
internal static partial class EwsUtilities
{
    /// <summary>
    ///     Gets the build version.
    /// </summary>
    /// <value>The build version.</value>
    internal const string BuildVersion = "15.0.913.15";


    internal const string XsFalse = "false";
    internal const string XsTrue = "true";

    internal const string EwsTypesNamespacePrefix = "t";
    internal const string EwsMessagesNamespacePrefix = "m";
    internal const string EwsErrorsNamespacePrefix = "e";
    internal const string EwsSoapNamespacePrefix = "soap";
    internal const string EwsXmlSchemaInstanceNamespacePrefix = "xsi";
    internal const string PassportSoapFaultNamespacePrefix = "psf";
    internal const string WsTrustFebruary2005NamespacePrefix = "wst";
    internal const string WsAddressingNamespacePrefix = "wsa";
    internal const string AutodiscoverSoapNamespacePrefix = "a";
    internal const string WsSecurityUtilityNamespacePrefix = "wsu";
    internal const string WsSecuritySecExtNamespacePrefix = "wsse";

    internal const string EwsTypesNamespace = "http://schemas.microsoft.com/exchange/services/2006/types";
    internal const string EwsMessagesNamespace = "http://schemas.microsoft.com/exchange/services/2006/messages";
    internal const string EwsErrorsNamespace = "http://schemas.microsoft.com/exchange/services/2006/errors";
    internal const string EwsSoapNamespace = "http://schemas.xmlsoap.org/soap/envelope/";
    internal const string EwsSoap12Namespace = "http://www.w3.org/2003/05/soap-envelope";
    internal const string EwsXmlSchemaInstanceNamespace = "http://www.w3.org/2001/XMLSchema-instance";
    internal const string PassportSoapFaultNamespace = "http://schemas.microsoft.com/Passport/SoapServices/SOAPFault";
    internal const string WsTrustFebruary2005Namespace = "http://schemas.xmlsoap.org/ws/2005/02/trust";

    internal const string WsAddressingNamespace = "http://www.w3.org/2005/08/addressing";

    internal const string AutodiscoverSoapNamespace = "http://schemas.microsoft.com/exchange/2010/Autodiscover";

    internal const string WsSecurityUtilityNamespace =
        "http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd";

    internal const string WsSecuritySecExtNamespace =
        "http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd";

    /// <summary>
    ///     Map from XML element names to ServiceObject type and constructors.
    /// </summary>
    private static readonly ServiceObjectInfo ServiceObjectInfo = new();


    /// <summary>
    ///     Dictionary of enum type to ExchangeVersion maps.
    /// </summary>
    private static readonly IReadOnlyDictionary<Type, Dictionary<Enum, ExchangeVersion>> EnumVersionDictionaries =
        new Dictionary<Type, Dictionary<Enum, ExchangeVersion>>
        {
            // @formatter:off
            { typeof(WellKnownFolderName), BuildEnumDict(typeof(WellKnownFolderName)) },
            { typeof(ItemTraversal), BuildEnumDict(typeof(ItemTraversal)) },
            { typeof(ConversationQueryTraversal), BuildEnumDict(typeof(ConversationQueryTraversal)) },
            { typeof(FileAsMapping), BuildEnumDict(typeof(FileAsMapping)) },
            { typeof(EventType), BuildEnumDict(typeof(EventType)) },
            { typeof(MeetingRequestsDeliveryScope), BuildEnumDict(typeof(MeetingRequestsDeliveryScope)) },
            { typeof(ViewFilter), BuildEnumDict(typeof(ViewFilter)) },
            // @formatter:on
        };

    /// <summary>
    ///     Dictionary of enum type to schema-name-to-enum-value maps.
    /// </summary>
    private static readonly IReadOnlyDictionary<Type, Dictionary<string, Enum>> SchemaToEnumDictionaries =
        new Dictionary<Type, Dictionary<string, Enum>>
        {
            // @formatter:off
            { typeof(EventType), BuildSchemaToEnumDict(typeof(EventType)) },
            { typeof(MailboxType), BuildSchemaToEnumDict(typeof(MailboxType)) },
            { typeof(FileAsMapping), BuildSchemaToEnumDict(typeof(FileAsMapping)) },
            { typeof(RuleProperty), BuildSchemaToEnumDict(typeof(RuleProperty)) },
            { typeof(WellKnownFolderName), BuildSchemaToEnumDict(typeof(WellKnownFolderName)) },
            // @formatter:on
        };

    /// <summary>
    ///     Dictionary of enum type to enum-value-to-schema-name maps.
    /// </summary>
    private static readonly IReadOnlyDictionary<Type, Dictionary<Enum, string>> EnumToSchemaDictionaries =
        new Dictionary<Type, Dictionary<Enum, string>>
        {
            // @formatter:off
            { typeof(EventType), BuildEnumToSchemaDict(typeof(EventType)) },
            { typeof(MailboxType), BuildEnumToSchemaDict(typeof(MailboxType)) },
            { typeof(FileAsMapping), BuildEnumToSchemaDict(typeof(FileAsMapping)) },
            { typeof(RuleProperty), BuildEnumToSchemaDict(typeof(RuleProperty)) },
            { typeof(WellKnownFolderName), BuildEnumToSchemaDict(typeof(WellKnownFolderName)) },
            // @formatter:on
        };

    /// <summary>
    ///     Dictionary to map from special CLR type names to their "short" names.
    /// </summary>
    private static readonly IReadOnlyDictionary<string, string> TypeNameToShortNameMap = new Dictionary<string, string>
    {
        // @formatter:off
        { "Boolean", "bool" },
        { "Int16", "short" },
        { "Int32", "int" },
        { "String", "string" },
        // @formatter:on
    };


    /// <summary>
    ///     Asserts that the specified condition if true.
    /// </summary>
    /// <param name="condition">Assertion.</param>
    /// <param name="caller">The caller.</param>
    /// <param name="message">The message to use if assertion fails.</param>
    [Conditional("DEBUG")]
    internal static void Assert([DoesNotReturnIf(false)] bool condition, string caller, string message)
    {
        Debug.Assert(condition, $"[{caller}] {message}");
    }

    /// <summary>
    ///     Gets the namespace prefix from an XmlNamespace enum value.
    /// </summary>
    /// <param name="xmlNamespace">The XML namespace.</param>
    /// <returns>Namespace prefix string.</returns>
    internal static string GetNamespacePrefix(XmlNamespace xmlNamespace)
    {
        return xmlNamespace switch
        {
            XmlNamespace.Types => EwsTypesNamespacePrefix,
            XmlNamespace.Messages => EwsMessagesNamespacePrefix,
            XmlNamespace.Errors => EwsErrorsNamespacePrefix,
            XmlNamespace.Soap => EwsSoapNamespacePrefix,
            XmlNamespace.Soap12 => EwsSoapNamespacePrefix,
            XmlNamespace.XmlSchemaInstance => EwsXmlSchemaInstanceNamespacePrefix,
            XmlNamespace.PassportSoapFault => PassportSoapFaultNamespacePrefix,
            XmlNamespace.WSTrustFebruary2005 => WsTrustFebruary2005NamespacePrefix,
            XmlNamespace.WSAddressing => WsAddressingNamespacePrefix,
            XmlNamespace.Autodiscover => AutodiscoverSoapNamespacePrefix,
            _ => string.Empty,
        };
    }

    /// <summary>
    ///     Gets the namespace URI from an XmlNamespace enum value.
    /// </summary>
    /// <param name="xmlNamespace">The XML namespace.</param>
    /// <returns>Uri as string</returns>
    internal static string GetNamespaceUri(XmlNamespace xmlNamespace)
    {
        return xmlNamespace switch
        {
            XmlNamespace.Types => EwsTypesNamespace,
            XmlNamespace.Messages => EwsMessagesNamespace,
            XmlNamespace.Errors => EwsErrorsNamespace,
            XmlNamespace.Soap => EwsSoapNamespace,
            XmlNamespace.Soap12 => EwsSoap12Namespace,
            XmlNamespace.XmlSchemaInstance => EwsXmlSchemaInstanceNamespace,
            XmlNamespace.PassportSoapFault => PassportSoapFaultNamespace,
            XmlNamespace.WSTrustFebruary2005 => WsTrustFebruary2005Namespace,
            XmlNamespace.WSAddressing => WsAddressingNamespace,
            XmlNamespace.Autodiscover => AutodiscoverSoapNamespace,
            _ => string.Empty,
        };
    }

    /// <summary>
    ///     Gets the XmlNamespace enum value from a namespace Uri.
    /// </summary>
    /// <param name="namespaceUri">XML namespace Uri.</param>
    /// <returns>XmlNamespace enum value.</returns>
    internal static XmlNamespace GetNamespaceFromUri(string namespaceUri)
    {
        return namespaceUri switch
        {
            EwsErrorsNamespace => XmlNamespace.Errors,
            EwsTypesNamespace => XmlNamespace.Types,
            EwsMessagesNamespace => XmlNamespace.Messages,
            EwsSoapNamespace => XmlNamespace.Soap,
            EwsSoap12Namespace => XmlNamespace.Soap12,
            EwsXmlSchemaInstanceNamespace => XmlNamespace.XmlSchemaInstance,
            PassportSoapFaultNamespace => XmlNamespace.PassportSoapFault,
            WsTrustFebruary2005Namespace => XmlNamespace.WSTrustFebruary2005,
            WsAddressingNamespace => XmlNamespace.WSAddressing,
            _ => XmlNamespace.NotSpecified,
        };
    }

    /// <summary>
    ///     Creates EWS object based on XML element name.
    /// </summary>
    /// <typeparam name="TServiceObject">The type of the service object.</typeparam>
    /// <param name="service">The service.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <returns>Service object.</returns>
    internal static TServiceObject? CreateEwsObjectFromXmlElementName<TServiceObject>(
        ExchangeService service,
        string xmlElementName
    )
        where TServiceObject : ServiceObject
    {
        if (ServiceObjectInfo.XmlElementNameToServiceObjectClassMap.TryGetValue(xmlElementName, out var itemClass))
        {
            if (ServiceObjectInfo.ServiceObjectConstructorsWithServiceParam.TryGetValue(
                    itemClass,
                    out var creationDelegate
                ))
            {
                return (TServiceObject)creationDelegate(service);
            }

            throw new ArgumentException(Strings.NoAppropriateConstructorForItemClass);
        }

        return default;
    }

    /// <summary>
    ///     Creates Item from Item class.
    /// </summary>
    /// <param name="itemAttachment">The item attachment.</param>
    /// <param name="itemClass">The item class.</param>
    /// <param name="isNew">If true, item attachment is new.</param>
    /// <returns>New Item.</returns>
    internal static Item CreateItemFromItemClass(ItemAttachment itemAttachment, Type itemClass, bool isNew)
    {
        if (ServiceObjectInfo.ServiceObjectConstructorsWithAttachmentParam.TryGetValue(
                itemClass,
                out var creationDelegate
            ))
        {
            return (Item)creationDelegate(itemAttachment, isNew);
        }

        throw new ArgumentException(Strings.NoAppropriateConstructorForItemClass);
    }

    /// <summary>
    ///     Creates Item based on XML element name.
    /// </summary>
    /// <param name="itemAttachment">The item attachment.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <returns>New Item.</returns>
    internal static Item? CreateItemFromXmlElementName(ItemAttachment itemAttachment, string xmlElementName)
    {
        if (ServiceObjectInfo.XmlElementNameToServiceObjectClassMap.TryGetValue(xmlElementName, out var itemClass))
        {
            return CreateItemFromItemClass(itemAttachment, itemClass, false);
        }

        return null;
    }

    /// <summary>
    ///     Gets the expected item type based on the local name.
    /// </summary>
    /// <param name="xmlElementName"></param>
    /// <returns></returns>
    internal static Type? GetItemTypeFromXmlElementName(string xmlElementName)
    {
        ServiceObjectInfo.XmlElementNameToServiceObjectClassMap.TryGetValue(xmlElementName, out var itemClass);
        return itemClass;
    }

    /// <summary>
    ///     Finds the first item of type TItem (not a descendant type) in the specified collection.
    /// </summary>
    /// <typeparam name="TItem">The type of the item to find.</typeparam>
    /// <param name="items">The collection.</param>
    /// <returns>A TItem instance or null if no instance of TItem could be found.</returns>
    internal static TItem? FindFirstItemOfType<TItem>(IEnumerable<Item> items)
        where TItem : Item
    {
        var itemType = typeof(TItem);

        foreach (var item in items)
        {
            // We're looking for an exact class match here.
            if (item.GetType() == itemType)
            {
                return (TItem)item;
            }
        }

        return null;
    }


    #region EmailAddress parsing

    /// <summary>
    ///     Gets the domain name from an email address.
    /// </summary>
    /// <param name="emailAddress">The email address.</param>
    /// <returns>Domain name.</returns>
    internal static string DomainFromEmailAddress(string emailAddress)
    {
        var emailAddressParts = emailAddress.Split('@');

        if (emailAddressParts.Length != 2 || string.IsNullOrEmpty(emailAddressParts[1]))
        {
            throw new FormatException(Strings.InvalidEmailAddress);
        }

        return emailAddressParts[1];
    }

    #endregion


    #region Tracing routines

    /// <summary>
    ///     Write trace start element.
    /// </summary>
    /// <param name="writer">The writer to write the start element to.</param>
    /// <param name="traceTag">The trace tag.</param>
    /// <param name="includeVersion">If true, include build version attribute.</param>
    private static void WriteTraceStartElement(XmlWriter writer, string traceTag, bool includeVersion)
    {
        writer.WriteStartElement("Trace");
        writer.WriteAttributeString("Tag", traceTag);
        writer.WriteAttributeString("Tid", Environment.CurrentManagedThreadId.ToString());
        writer.WriteAttributeString("Time", DateTime.UtcNow.ToString("u", DateTimeFormatInfo.InvariantInfo));

        if (includeVersion)
        {
            writer.WriteAttributeString("Version", BuildVersion);
        }
    }

    /// <summary>
    ///     Format log message.
    /// </summary>
    /// <param name="entryKind">Kind of the entry.</param>
    /// <param name="logEntry">The log entry.</param>
    /// <returns>XML log entry as a string.</returns>
    internal static string FormatLogMessage(string entryKind, string logEntry)
    {
        var sb = new StringBuilder();

        using (var writer = new StringWriter(sb))
        {
            using var xmlWriter = XmlWriter.Create(
                writer,
                new XmlWriterSettings
                {
                    Indent = true,
                }
            );
            WriteTraceStartElement(xmlWriter, entryKind, false);

            xmlWriter.WriteWhitespace(Environment.NewLine);
            xmlWriter.WriteValue(logEntry);
            xmlWriter.WriteWhitespace(Environment.NewLine);

            xmlWriter.WriteEndElement(); // Trace
            xmlWriter.WriteWhitespace(Environment.NewLine);
        }

        return sb.ToString();
    }

    /// <summary>
    ///     Format the HTTP headers.
    /// </summary>
    /// <param name="sb">StringBuilder.</param>
    /// <param name="headers">The HTTP headers.</param>
    private static void FormatHttpHeaders(StringBuilder sb, HttpHeaders headers)
    {
        foreach (var item in headers)
        {
            foreach (var value in item.Value)
            {
                sb.Append($"{item.Key}: {value}\n");
            }
        }
    }

    /// <summary>
    ///     Format request HTTP headers.
    /// </summary>
    /// <param name="request">The HTTP request.</param>
    internal static string FormatHttpRequestHeaders(EwsHttpWebRequest request)
    {
        var sb = new StringBuilder();
        sb.Append($"{request.Method} {request.RequestUri.AbsolutePath} HTTP/1.1\n");
        FormatHttpHeaders(sb, request.Headers);
        sb.Append('\n');

        return sb.ToString();
    }

    /// <summary>
    ///     Format response HTTP headers.
    /// </summary>
    /// <param name="response">The HTTP response.</param>
    internal static string FormatHttpResponseHeaders(IEwsHttpWebResponse response)
    {
        var sb = new StringBuilder();
        sb.Append($"HTTP/{response.ProtocolVersion} {(int)response.StatusCode} {response.StatusDescription}\n");

        sb.Append(FormatHttpHeaders(response.Headers));
        sb.Append('\n');
        return sb.ToString();
    }

    /// <summary>
    ///     Formats HTTP headers.
    /// </summary>
    /// <param name="headers">The headers.</param>
    /// <returns>Headers as a string</returns>
    private static string FormatHttpHeaders(HttpHeaders headers)
    {
        var sb = new StringBuilder();
        foreach (var item in headers)
        {
            foreach (var value in item.Value)
            {
                sb.Append($"{item.Key}: {value}\n");
            }
        }

        return sb.ToString();
    }

    /// <summary>
    ///     Format XML content in a MemoryStream for message.
    /// </summary>
    /// <param name="entryKind">Kind of the entry.</param>
    /// <param name="memoryStream">The memory stream.</param>
    /// <returns>XML log entry as a string.</returns>
    internal static string FormatLogMessageWithXmlContent(string entryKind, MemoryStream memoryStream)
    {
        var sb = new StringBuilder();
        var settings = new XmlReaderSettings
        {
            ConformanceLevel = ConformanceLevel.Fragment,
            IgnoreComments = true,
            IgnoreWhitespace = true,
            CloseInput = false,
        };


        // Remember the current location in the MemoryStream.
        var lastPosition = memoryStream.Position;

        // Rewind the position since we want to format the entire contents.
        memoryStream.Position = 0;

        try
        {
            using var reader = XmlReader.Create(memoryStream, settings);
            reader.Read();
            if (reader.NodeType == XmlNodeType.XmlDeclaration)
            {
                reader.Read();
            }

            using var writer = new StringWriter(sb);
            using var xmlWriter = XmlWriter.Create(
                writer,
                new XmlWriterSettings
                {
                    Indent = true,
                }
            );
            WriteTraceStartElement(xmlWriter, entryKind, true);

            while (!reader.EOF)
            {
                xmlWriter.WriteNode(reader, true);
            }

            xmlWriter.WriteEndElement(); // Trace
            xmlWriter.WriteWhitespace(Environment.NewLine);
        }
        catch (XmlException)
        {
            // We tried to format the content as "pretty" XML. Apparently the content is
            // not well-formed XML or isn't XML at all. Fallback and treat it as plain text.
            sb.Length = 0;
            memoryStream.Position = 0;
            sb.Append(memoryStream);
        }

        // Restore Position in the stream.
        memoryStream.Position = lastPosition;

        return sb.ToString();
    }

    #endregion


    #region Conversion routines

    /// <summary>
    ///     Convert bool to XML Schema bool.
    /// </summary>
    /// <param name="value">Bool value.</param>
    /// <returns>String representing bool value in XML Schema.</returns>
    internal static string BoolToXsBool(bool value)
    {
        return value ? XsTrue : XsFalse;
    }

    /// <summary>
    ///     Parses an enum value list.
    /// </summary>
    /// <typeparam name="T">Type of value.</typeparam>
    /// <param name="list">The list.</param>
    /// <param name="value">The value.</param>
    /// <param name="separators">The separators.</param>
    internal static void ParseEnumValueList<T>(IList<T> list, string value, params char[] separators)
        where T : struct
    {
        Assert(typeof(T).GetTypeInfo().IsEnum, "EwsUtilities.ParseEnumValueList", "T is not an enum type.");

        if (string.IsNullOrEmpty(value))
        {
            return;
        }

        var enumValues = value.Split(separators);

        foreach (var enumValue in enumValues)
        {
            list.Add(Enum.Parse<T>(enumValue, ignoreCase: false));
        }
    }

    /// <summary>
    ///     Converts an enum to a string, using the mapping dictionaries if appropriate.
    /// </summary>
    /// <param name="value">The enum value to be serialized</param>
    /// <returns>String representation of enum to be used in the protocol</returns>
    internal static string SerializeEnum(Enum value)
    {
        if (EnumToSchemaDictionaries.TryGetValue(value.GetType(), out var enumToStringDict) &&
            enumToStringDict.TryGetValue(value, out var strValue))
        {
            return strValue;
        }

        return value.ToString();
    }

    /// <summary>
    ///     Parses specified value based on type.
    /// </summary>
    /// <typeparam name="T">Type of value.</typeparam>
    /// <param name="value">The value.</param>
    /// <returns>Value of type T.</returns>
    internal static T Parse<T>(string value)
    {
        if (typeof(T).GetTypeInfo().IsEnum)
        {
            if (SchemaToEnumDictionaries.TryGetValue(typeof(T), out var stringToEnumDict) &&
                stringToEnumDict.TryGetValue(value, out var enumValue))
            {
                // This double-casting is ugly, but necessary. By this point, we know that T is an Enum
                // (same as returned by the dictionary), but the compiler can't prove it. Thus, the 
                // up-cast before we can down-cast.
                return (T)(object)enumValue;
            }

            return (T)Enum.Parse(typeof(T), value, ignoreCase: false);
        }

        return (T)Convert.ChangeType(value, typeof(T), CultureInfo.InvariantCulture);
    }

    /// <summary>
    ///     Tries to parses the specified value to the specified type.
    /// </summary>
    /// <typeparam name="T">The type into which to cast the provided value.</typeparam>
    /// <param name="value">The value to parse.</param>
    /// <param name="result">
    ///     The value cast to the specified type, if TryParse succeeds. Otherwise, the value of result is
    ///     indeterminate.
    /// </param>
    /// <returns>True if value could be parsed; otherwise, false.</returns>
    internal static bool TryParse<T>(string value, [MaybeNullWhen(false)] out T result)
    {
        try
        {
            result = Parse<T>(value);
            return true;
        }
        //// Catch all exceptions here, we're not interested in the reason why TryParse failed.
        catch (Exception)
        {
            result = default;
            return false;
        }
    }

    /// <summary>
    ///     Converts the specified date and time from one time zone to another.
    /// </summary>
    /// <param name="dateTime">The date time to convert.</param>
    /// <param name="sourceTimeZone">The source time zone.</param>
    /// <param name="destinationTimeZone">The destination time zone.</param>
    /// <returns>A DateTime that holds the converted</returns>
    internal static DateTime ConvertTime(
        DateTime dateTime,
        TimeZoneInfo sourceTimeZone,
        TimeZoneInfo destinationTimeZone
    )
    {
        try
        {
            return TimeZoneInfo.ConvertTime(dateTime, sourceTimeZone, destinationTimeZone);
        }
        catch (ArgumentException e)
        {
            throw new TimeZoneConversionException(
                string.Format(
                    Strings.CannotConvertBetweenTimeZones,
                    DateTimeToXsDateTime(dateTime),
                    sourceTimeZone.DisplayName,
                    destinationTimeZone.DisplayName
                ),
                e
            );
        }
    }

    /// <summary>
    ///     Reads the string as date time, assuming it is unbiased (e.g. 2009/01/01T08:00)
    ///     and scoped to service's time zone.
    /// </summary>
    /// <param name="dateString">The date string.</param>
    /// <param name="service">The service.</param>
    /// <returns>The string's value as a DateTime object.</returns>
    internal static DateTime ParseAsUnbiasedDatetimescopedToServicetimeZone(string dateString, ExchangeService service)
    {
        // Convert the element's value to a DateTime with no adjustment.
        var tempDate = DateTime.Parse(dateString, CultureInfo.InvariantCulture);

        // Set the kind according to the service's time zone
        if (service.TimeZone.Equals(TimeZoneInfo.Utc))
        {
            return new DateTime(tempDate.Ticks, DateTimeKind.Utc);
        }

        if (IsLocalTimeZone(service.TimeZone))
        {
            return new DateTime(tempDate.Ticks, DateTimeKind.Local);
        }

        return new DateTime(tempDate.Ticks, DateTimeKind.Unspecified);
    }

    /// <summary>
    ///     Determines whether the specified time zone is the same as the system's local time zone.
    /// </summary>
    /// <param name="timeZone">The time zone to check.</param>
    /// <returns>
    ///     <c>true</c> if the specified time zone is the same as the system's local time zone; otherwise, <c>false</c>.
    /// </returns>
    internal static bool IsLocalTimeZone(TimeZoneInfo timeZone)
    {
        return TimeZoneInfo.Local.Equals(timeZone) ||
               TimeZoneInfo.Local.Id == timeZone.Id /* && TimeZoneInfo.Local.HasSameRules(timeZone)*/;
    }

    /// <summary>
    ///     Convert DateTime to XML Schema date.
    /// </summary>
    /// <param name="date">The date to be converted.</param>
    /// <returns>String representation of DateTime.</returns>
    internal static string DateTimeToXsDate(DateTime date)
    {
        // Depending on the current culture, DateTime formatter will 
        // translate dates from one culture to another (e.g. Gregorian to Lunar).  The server
        // however, considers all dates to be in Gregorian, so using the InvariantCulture will
        // ensure this.

        var format = date.Kind switch
        {
            DateTimeKind.Utc => "yyyy-MM-ddZ",
            DateTimeKind.Unspecified => "yyyy-MM-dd",
            // DateTimeKind.Local is remaining
            _ => "yyyy-MM-ddzzz",
        };

        return date.ToString(format, CultureInfo.InvariantCulture);
    }

    /// <summary>
    ///     Dates the DateTime into an XML schema date time.
    /// </summary>
    /// <param name="dateTime">The date time.</param>
    /// <returns>String representation of DateTime.</returns>
    internal static string DateTimeToXsDateTime(DateTime dateTime)
    {
        var format = "yyyy-MM-ddTHH:mm:ss.fff";

        switch (dateTime.Kind)
        {
            case DateTimeKind.Utc:
            {
                format += "Z";
                break;
            }
            case DateTimeKind.Local:
            {
                format += "zzz";
                break;
            }
        }

        // Depending on the current culture, DateTime formatter will replace ':' with 
        // the DateTimeFormatInfo.TimeSeparator property which may not be ':'. Force the proper string
        // to be used by using the InvariantCulture.
        return dateTime.ToString(format, CultureInfo.InvariantCulture);
    }

    /// <summary>
    ///     Convert EWS DayOfTheWeek enum to System.DayOfWeek.
    /// </summary>
    /// <param name="dayOfTheWeek">The day of the week.</param>
    /// <returns>System.DayOfWeek value.</returns>
    internal static DayOfWeek EwsToSystemDayOfWeek(DayOfTheWeek dayOfTheWeek)
    {
        if (dayOfTheWeek == DayOfTheWeek.Day ||
            dayOfTheWeek == DayOfTheWeek.Weekday ||
            dayOfTheWeek == DayOfTheWeek.WeekendDay)
        {
            throw new ArgumentException(
                $"Cannot convert {dayOfTheWeek} to System.DayOfWeek enum value",
                nameof(dayOfTheWeek)
            );
        }

        return (DayOfWeek)dayOfTheWeek;
    }

    /// <summary>
    ///     Convert System.DayOfWeek type to EWS DayOfTheWeek.
    /// </summary>
    /// <param name="dayOfWeek">The dayOfWeek.</param>
    /// <returns>EWS DayOfWeek value</returns>
    internal static DayOfTheWeek SystemToEwsDayOfTheWeek(DayOfWeek dayOfWeek)
    {
        return (DayOfTheWeek)dayOfWeek;
    }

    /// <summary>
    ///     Takes a System.TimeSpan structure and converts it into an
    ///     xs:duration string as defined by the W3 Consortiums Recommendation
    ///     "XML Schema Part 2: Datatypes Second Edition",
    ///     http://www.w3.org/TR/xmlschema-2/#duration
    /// </summary>
    /// <param name="timeSpan">TimeSpan structure to convert</param>
    /// <returns>xs:duration formatted string</returns>
    internal static string TimeSpanToXsDuration(TimeSpan timeSpan)
    {
        // Optional '-' offset
        var offsetStr = timeSpan.TotalSeconds < 0 ? "-" : string.Empty;

        // The TimeSpan structure does not have a Year or Month 
        // property, therefore we wouldn't be able to return an xs:duration
        // string from a TimeSpan that included the nY or nM components.
        return string.Format(
            "{0}P{1}DT{2}H{3}M{4}S",
            offsetStr,
            Math.Abs(timeSpan.Days),
            Math.Abs(timeSpan.Hours),
            Math.Abs(timeSpan.Minutes),
            Math.Abs(timeSpan.Seconds) + "." + Math.Abs(timeSpan.Milliseconds)
        );
    }

    /// <summary>
    ///     Takes an xs:duration string as defined by the W3 Consortiums
    ///     Recommendation "XML Schema Part 2: Datatypes Second Edition",
    ///     http://www.w3.org/TR/xmlschema-2/#duration, and converts it
    ///     into a System.TimeSpan structure
    /// </summary>
    /// <remarks>
    ///     This method uses the following approximations:
    ///     1 year = 365 days
    ///     1 month = 30 days
    ///     Additionally, it only allows for four decimal points of
    ///     seconds precision.
    /// </remarks>
    /// <param name="xsDuration">xs:duration string to convert</param>
    /// <returns>System.TimeSpan structure</returns>
    internal static TimeSpan XsDurationToTimeSpan(string xsDuration)
    {
        var timeSpanParser = TimeSpanParserRegex();

        var m = timeSpanParser.Match(xsDuration);
        if (!m.Success)
        {
            throw new ArgumentException(Strings.XsDurationCouldNotBeParsed);
        }

        var token = m.Result("${pos}");
        var negative = !string.IsNullOrEmpty(token);

        // Year
        token = m.Result("${year}");
        var year = 0;
        if (!string.IsNullOrEmpty(token))
        {
            year = int.Parse(token);
        }

        // Month
        token = m.Result("${month}");
        var month = 0;
        if (!string.IsNullOrEmpty(token))
        {
            month = int.Parse(token);
        }

        // Day
        token = m.Result("${day}");
        var day = 0;
        if (!string.IsNullOrEmpty(token))
        {
            day = int.Parse(token);
        }

        // Hour
        token = m.Result("${hour}");
        var hour = 0;
        if (!string.IsNullOrEmpty(token))
        {
            hour = int.Parse(token);
        }

        // Minute
        token = m.Result("${minute}");
        var minute = 0;
        if (!string.IsNullOrEmpty(token))
        {
            minute = int.Parse(token);
        }

        // Seconds
        token = m.Result("${seconds}");
        var seconds = 0;
        if (!string.IsNullOrEmpty(token))
        {
            seconds = int.Parse(token);
        }

        var milliseconds = 0;
        token = m.Result("${precision}");

        // Only allowed 4 digits of precision
        if (token.Length > 4)
        {
            token = token.Substring(0, 4);
        }

        if (!string.IsNullOrEmpty(token))
        {
            milliseconds = int.Parse(token);
        }

        // Apply conversions of year and months to days.
        // Year = 365 days
        // Month = 30 days
        day = day + year * 365 + month * 30;
        var retVal = new TimeSpan(day, hour, minute, seconds, milliseconds);

        if (negative)
        {
            retVal = -retVal;
        }

        return retVal;
    }

    /// <summary>
    ///     Converts the specified time span to its XSD representation.
    /// </summary>
    /// <param name="timeSpan">The time span.</param>
    /// <returns>The XSD representation of the specified time span.</returns>
    public static string TimeSpanToXsTime(TimeSpan timeSpan)
    {
        return string.Format("{0:00}:{1:00}:{2:00}", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
    }

    #endregion


    #region Type Name utilities

    /// <summary>
    ///     Gets the printable name of a CLR type.
    /// </summary>
    /// <param name="type">The type.</param>
    /// <returns>Printable name.</returns>
    public static string GetPrintableTypeName(Type type)
    {
        if (type.GetTypeInfo().IsGenericType)
        {
            // Convert generic type to printable form (e.g. List<Item>)
            var genericPrefix = type.Name.Substring(0, type.Name.IndexOf('`'));
            var nameBuilder = new StringBuilder(genericPrefix);

            // Note: building array of generic parameters is done recursively. Each parameter could be any type.
            var genericArgs = type.GetGenericArguments().ToList().Select(GetPrintableTypeName).ToArray();

            nameBuilder.Append('<');
            nameBuilder.Append(string.Join(",", genericArgs));
            nameBuilder.Append('>');
            return nameBuilder.ToString();
        }

        if (type.IsArray)
        {
            // Convert array type to printable form.
            var arrayPrefix = type.Name.Substring(0, type.Name.IndexOf('['));
            var nameBuilder = new StringBuilder(GetSimplifiedTypeName(arrayPrefix));
            for (var rank = 0; rank < type.GetArrayRank(); rank++)
            {
                nameBuilder.Append("[]");
            }

            return nameBuilder.ToString();
        }

        return GetSimplifiedTypeName(type.Name);
    }

    /// <summary>
    ///     Gets the printable name of a simple CLR type.
    /// </summary>
    /// <param name="typeName">The type name.</param>
    /// <returns>Printable name.</returns>
    private static string GetSimplifiedTypeName(string typeName)
    {
        // If type has a shortname (e.g. int for Int32) map to the short name.
        return TypeNameToShortNameMap.TryGetValue(typeName, out var name) ? name : typeName;
    }

    #endregion


    #region Method parameters validation routines

    /// <summary>
    ///     Validates parameter (and allows null value).
    /// </summary>
    /// <param name="param">The param.</param>
    /// <param name="paramName">Name of the param.</param>
    internal static void ValidateParamAllowNull(
        [NotNull] object? param,
        [CallerArgumentExpression(nameof(param))] string? paramName = null
    )
    {
        switch (param)
        {
            case ISelfValidate selfValidate:
            {
                try
                {
                    selfValidate.Validate();
                }
                catch (ServiceValidationException e)
                {
                    throw new ArgumentException(Strings.ValidationFailed, paramName, e);
                }

                break;
            }
            case ServiceObject ewsObject when ewsObject.IsNew:
            {
                throw new ArgumentException(Strings.ObjectDoesNotHaveId, paramName);
            }
        }
    }

    /// <summary>
    ///     Validates parameter (null value not allowed).
    /// </summary>
    /// <param name="param">The param.</param>
    /// <param name="paramName">Name of the param.</param>
    internal static void ValidateParam(
        [NotNull] object? param,
        [CallerArgumentExpression(nameof(param))] string? paramName = null
    )
    {
        bool isValid;

        if (param is string strParam)
        {
            isValid = !string.IsNullOrEmpty(strParam);
        }
        else
        {
            isValid = param != null;
        }

        if (!isValid)
        {
            throw new ArgumentNullException(paramName);
        }

        ValidateParamAllowNull(param, paramName);
    }

    /// <summary>
    ///     Validates parameter collection.
    /// </summary>
    /// <param name="collection">The collection.</param>
    /// <param name="paramName">Name of the param.</param>
    internal static void ValidateParamCollection(
        IEnumerable collection,
        [CallerArgumentExpression(nameof(collection))] string? paramName = null
    )
    {
        ValidateParam(collection, paramName);

        var count = 0;

        foreach (var obj in collection)
        {
            try
            {
                ValidateParam(obj, $"collection[{count}]");
            }
            catch (ArgumentException e)
            {
                throw new ArgumentException($"The element at position {count} is invalid", paramName, e);
            }

            count++;
        }

        if (count == 0)
        {
            throw new ArgumentException(Strings.CollectionIsEmpty, paramName);
        }
    }

    /// <summary>
    ///     Validates string parameter to be non-empty string (null value allowed).
    /// </summary>
    /// <param name="param">The string parameter.</param>
    /// <param name="paramName">Name of the parameter.</param>
    internal static void ValidateNonBlankStringParamAllowNull(
        string? param,
        [CallerArgumentExpression(nameof(param))] string? paramName = null
    )
    {
        if (param != null)
        {
            // Non-empty string has at least one character which is *not* a whitespace character
            if (param.Length == param.CountMatchingChars(char.IsWhiteSpace))
            {
                throw new ArgumentException(Strings.ArgumentIsBlankString, paramName);
            }
        }
    }

    /// <summary>
    ///     Validates string parameter to be non-empty string (null value not allowed).
    /// </summary>
    /// <param name="param">The string parameter.</param>
    /// <param name="paramName">Name of the parameter.</param>
    internal static void ValidateNonBlankStringParam(string param, string paramName)
    {
        if (param == null)
        {
            throw new ArgumentNullException(paramName);
        }

        ValidateNonBlankStringParamAllowNull(param, paramName);
    }

    /// <summary>
    ///     Validates the enum value against the request version.
    /// </summary>
    /// <param name="enumValue">The enum value.</param>
    /// <param name="requestVersion">The request version.</param>
    /// <exception cref="ServiceVersionException">Raised if this enum value requires a later version of Exchange.</exception>
    internal static void ValidateEnumVersionValue(Enum enumValue, ExchangeVersion requestVersion)
    {
        var enumType = enumValue.GetType();
        var enumVersionDict = EnumVersionDictionaries[enumType];
        var enumVersion = enumVersionDict[enumValue];

        if (requestVersion < enumVersion)
        {
            throw new ServiceVersionException(
                string.Format(
                    Strings.EnumValueIncompatibleWithRequestVersion,
                    enumValue.ToString(),
                    enumType.Name,
                    enumVersion
                )
            );
        }
    }

    /// <summary>
    ///     Validates service object version against the request version.
    /// </summary>
    /// <param name="serviceObject">The service object.</param>
    /// <param name="requestVersion">The request version.</param>
    /// <exception cref="ServiceVersionException">Raised if this service object type requires a later version of Exchange.</exception>
    internal static void ValidateServiceObjectVersion(ServiceObject serviceObject, ExchangeVersion requestVersion)
    {
        var minimumRequiredServerVersion = serviceObject.GetMinimumRequiredServerVersion();

        if (requestVersion < minimumRequiredServerVersion)
        {
            throw new ServiceVersionException(
                string.Format(
                    Strings.ObjectTypeIncompatibleWithRequestVersion,
                    serviceObject.GetType().Name,
                    minimumRequiredServerVersion
                )
            );
        }
    }

    /// <summary>
    ///     Validates property version against the request version.
    /// </summary>
    /// <param name="service">The Exchange service.</param>
    /// <param name="minimumServerVersion">The minimum server version that supports the property.</param>
    /// <param name="propertyName">Name of the property.</param>
    internal static void ValidatePropertyVersion(
        ExchangeService service,
        ExchangeVersion minimumServerVersion,
        string propertyName
    )
    {
        if (service.RequestedServerVersion < minimumServerVersion)
        {
            throw new ServiceVersionException(
                string.Format(Strings.PropertyIncompatibleWithRequestVersion, propertyName, minimumServerVersion)
            );
        }
    }

    /// <summary>
    ///     Validates method version against the request version.
    /// </summary>
    /// <param name="service">The Exchange service.</param>
    /// <param name="minimumServerVersion">The minimum server version that supports the method.</param>
    /// <param name="methodName">Name of the method.</param>
    internal static void ValidateMethodVersion(
        ExchangeService service,
        ExchangeVersion minimumServerVersion,
        string methodName
    )
    {
        if (service.RequestedServerVersion < minimumServerVersion)
        {
            throw new ServiceVersionException(
                string.Format(Strings.MethodIncompatibleWithRequestVersion, methodName, minimumServerVersion)
            );
        }
    }

    /// <summary>
    ///     Validates class version against the request version.
    /// </summary>
    /// <param name="service">The Exchange service.</param>
    /// <param name="minimumServerVersion">The minimum server version that supports the method.</param>
    /// <param name="className">Name of the class.</param>
    internal static void ValidateClassVersion(
        ExchangeService service,
        ExchangeVersion minimumServerVersion,
        string className
    )
    {
        if (service.RequestedServerVersion < minimumServerVersion)
        {
            throw new ServiceVersionException(
                string.Format(Strings.ClassIncompatibleWithRequestVersion, className, minimumServerVersion)
            );
        }
    }

    /// <summary>
    ///     Validates domain name (null value allowed)
    /// </summary>
    /// <param name="domainName">Domain name.</param>
    /// <param name="paramName">Parameter name.</param>
    internal static void ValidateDomainNameAllowNull(
        string? domainName,
        [CallerArgumentExpression(nameof(domainName))] string? paramName = null
    )
    {
        if (domainName != null)
        {
            var regex = DomainRegex();

            if (!regex.IsMatch(domainName))
            {
                throw new ArgumentException(string.Format(Strings.InvalidDomainName, domainName), paramName);
            }
        }
    }

    /// <summary>
    ///     Gets version for enum member.
    /// </summary>
    /// <param name="enumType">Type of the enum.</param>
    /// <param name="enumName">The enum name.</param>
    /// <returns>Exchange version in which the enum value was first defined.</returns>
    private static ExchangeVersion GetEnumVersion(
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicFields)] Type enumType,
        string enumName
    )
    {
        var fieldInfo = enumType.GetField(enumName);
        Assert(
            fieldInfo != null,
            "EwsUtilities.GetEnumVersion",
            "Enum member " + enumName + " not found in " + enumType
        );

        var attrs = fieldInfo.GetCustomAttributes<RequiredServerVersionAttribute>(false).ToArray();
        if (attrs != null && attrs.Length > 0)
        {
            return attrs[0].Version;
        }

        return ExchangeVersion.Exchange2007_SP1;
    }

    /// <summary>
    ///     Builds the enum to version mapping dictionary.
    /// </summary>
    /// <param name="enumType">Type of the enum.</param>
    /// <returns>Dictionary of enum values to versions.</returns>
    private static Dictionary<Enum, ExchangeVersion> BuildEnumDict(
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicFields)] Type enumType
    )
    {
        var dict = new Dictionary<Enum, ExchangeVersion>();

        var names = Enum.GetNames(enumType);
        foreach (var name in names)
        {
            var value = (Enum)Enum.Parse(enumType, name, false);
            var version = GetEnumVersion(enumType, name);
            dict.Add(value, version);
        }

        return dict;
    }

    /// <summary>
    ///     Gets the schema name for enum member.
    /// </summary>
    /// <param name="enumType">Type of the enum.</param>
    /// <param name="enumName">The enum name.</param>
    /// <returns>The name for the enum used in the protocol, or null if it is the same as the enum's ToString().</returns>
    private static string? GetEnumSchemaName(
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicFields)] Type enumType,
        string enumName
    )
    {
        var fieldInfo = enumType.GetField(enumName);
        Assert(
            fieldInfo != null,
            "EwsUtilities.GetEnumSchemaName",
            "Enum member " + enumName + " not found in " + enumType
        );

        var attrs = fieldInfo.GetCustomAttributes<EwsEnumAttribute>(false).ToArray();
        if (attrs != null && attrs.Length > 0)
        {
            return attrs[0].SchemaName;
        }

        return null;
    }

    /// <summary>
    ///     Builds the schema to enum mapping dictionary.
    /// </summary>
    /// <param name="enumType">Type of the enum.</param>
    /// <returns>The mapping from enum to schema name</returns>
    private static Dictionary<string, Enum> BuildSchemaToEnumDict(
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicFields)] Type enumType
    )
    {
        var dict = new Dictionary<string, Enum>();
        var names = Enum.GetNames(enumType);

        foreach (var name in names)
        {
            var value = (Enum)Enum.Parse(enumType, name, false);
            var schemaName = GetEnumSchemaName(enumType, name);

            if (!string.IsNullOrEmpty(schemaName))
            {
                dict.Add(schemaName, value);
            }
        }

        return dict;
    }

    /// <summary>
    ///     Builds the enum to schema mapping dictionary.
    /// </summary>
    /// <param name="enumType">Type of the enum.</param>
    /// <returns>The mapping from enum to schema name</returns>
    private static Dictionary<Enum, string> BuildEnumToSchemaDict(
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicFields)] Type enumType
    )
    {
        var dict = new Dictionary<Enum, string>();
        var names = Enum.GetNames(enumType);
        foreach (var name in names)
        {
            var value = (Enum)Enum.Parse(enumType, name, false);
            var schemaName = GetEnumSchemaName(enumType, name);

            if (!string.IsNullOrEmpty(schemaName))
            {
                dict.Add(value, schemaName);
            }
        }

        return dict;
    }

    #endregion


    #region IEnumerable utility methods

    /// <summary>
    ///     Gets the enumerated object count.
    /// </summary>
    /// <param name="objects">The objects.</param>
    /// <returns>Count of objects in IEnumerable.</returns>
    internal static int GetEnumeratedObjectCount(IEnumerable objects)
    {
        var count = 0;

        foreach (var obj in objects)
        {
            count++;
        }

        return count;
    }

    /// <summary>
    ///     Gets enumerated object at index.
    /// </summary>
    /// <param name="objects">The objects.</param>
    /// <param name="index">The index.</param>
    /// <returns>Object at index.</returns>
    internal static object GetEnumeratedObjectAt(IEnumerable objects, int index)
    {
        var count = 0;

        foreach (var obj in objects)
        {
            if (count == index)
            {
                return obj;
            }

            count++;
        }

        throw new ArgumentOutOfRangeException(nameof(index), Strings.IEnumerableDoesNotContainThatManyObject);
    }

    #endregion


    #region Extension methods

    /// <summary>
    ///     Count characters in string that match a condition.
    /// </summary>
    /// <param name="str">The string.</param>
    /// <param name="charPredicate">Predicate to evaluate for each character in the string.</param>
    /// <returns>Count of characters that match condition expressed by predicate.</returns>
    private static int CountMatchingChars(this string str, Predicate<char> charPredicate)
    {
        return str.Count(ch => charPredicate(ch));
    }

    /// <summary>
    ///     Determines whether every element in the collection matches the conditions defined by the specified predicate.
    /// </summary>
    /// <typeparam name="T">Entry type.</typeparam>
    /// <param name="collection">The collection.</param>
    /// <param name="predicate">Predicate that defines the conditions to check against the elements.</param>
    /// <returns>
    ///     True if every element in the collection matches the conditions defined by the specified predicate; otherwise,
    ///     false.
    /// </returns>
    internal static bool TrueForAll<T>(this IEnumerable<T> collection, Predicate<T> predicate)
    {
        foreach (var entry in collection)
        {
            if (!predicate(entry))
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    ///     Call an action for each member of a collection.
    /// </summary>
    /// <param name="collection">The collection.</param>
    /// <param name="action">The action to apply.</param>
    /// <typeparam name="T">Collection element type.</typeparam>
    internal static void ForEach<T>(this IEnumerable<T> collection, Action<T> action)
    {
        foreach (var entry in collection)
        {
            action(entry);
        }
    }


    /*
       "(?<pos>-)?" +
       "P" +
       "((?<year>[0-9]+)Y)?" +
       "((?<month>[0-9]+)M)?" +
       "((?<day>[0-9]+)D)?" +
       "(T" +
       "((?<hour>[0-9]+)H)?" +
       "((?<minute>[0-9]+)M)?" +
       "((?<seconds>[0-9]+)(\\.(?<precision>[0-9]+))?S)?)?"
     */
    [GeneratedRegex(
        "(?<pos>-)?P((?<year>[0-9]+)Y)?((?<month>[0-9]+)M)?((?<day>[0-9]+)D)?(T((?<hour>[0-9]+)H)?((?<minute>[0-9]+)M)?((?<seconds>[0-9]+)(\\.(?<precision>[0-9]+))?S)?)?"
    )]
    private static partial Regex TimeSpanParserRegex();

    /// <summary>
    ///     Regular expression for legal domain names.
    /// </summary>
    [GeneratedRegex("^[-a-zA-Z0-9_.]+$")]
    private static partial Regex DomainRegex();

    #endregion
}
