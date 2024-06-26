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
///     XML reader.
/// </summary>
internal class EwsServiceXmlReader : EwsXmlReader
{
    /// <summary>
    ///     Gets the service.
    /// </summary>
    /// <value>The service.</value>
    public ExchangeService Service { get; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="EwsServiceXmlReader" /> class.
    /// </summary>
    /// <param name="stream">The stream.</param>
    /// <param name="service">The service.</param>
    internal EwsServiceXmlReader(Stream stream, ExchangeService service)
        : base(stream)
    {
        Service = service;
    }


    /// <summary>
    ///     Converts the specified string into a DateTime objects.
    /// </summary>
    /// <param name="dateTimeString">The date time string to convert.</param>
    /// <returns>A DateTime representing the converted string.</returns>
    private DateTime? ConvertStringToDateTime(string? dateTimeString)
    {
        return Service.ConvertUniversalDateTimeStringToLocalDateTime(dateTimeString);
    }

    /// <summary>
    ///     Converts the specified string into a unspecified Date object, ignoring offset.
    /// </summary>
    /// <param name="dateTimeString">The date time string to convert.</param>
    /// <returns>A DateTime representing the converted string.</returns>
    private static DateTime? ConvertStringToUnspecifiedDate(string? dateTimeString)
    {
        return ExchangeServiceBase.ConvertStartDateToUnspecifiedDateTime(dateTimeString);
    }

    /// <summary>
    ///     Reads the element value as date time.
    /// </summary>
    /// <returns>Element value.</returns>
    public DateTime? ReadElementValueAsDateTime()
    {
        return ConvertStringToDateTime(ReadElementValue());
    }

    /// <summary>
    ///     Reads the element value as unspecified date.
    /// </summary>
    /// <returns>Element value.</returns>
    public DateTime? ReadElementValueAsUnspecifiedDate()
    {
        return ConvertStringToUnspecifiedDate(ReadElementValue());
    }

    /// <summary>
    ///     Reads the element value as date time, assuming it is unbiased (e.g. 2009/01/01T08:00)
    ///     and scoped to service's time zone.
    /// </summary>
    /// <returns>The element's value as a DateTime object.</returns>
    public DateTime ReadElementValueAsUnbiasedDateTimeScopedToServiceTimeZone()
    {
        var elementValue = ReadElementValue();
        return EwsUtilities.ParseAsUnbiasedDatetimescopedToServicetimeZone(elementValue, Service);
    }

    /// <summary>
    ///     Reads the element value as date time.
    /// </summary>
    /// <param name="xmlNamespace">The XML namespace.</param>
    /// <param name="localName">Name of the local.</param>
    /// <returns>Element value.</returns>
    public DateTime? ReadElementValueAsDateTime(XmlNamespace xmlNamespace, string localName)
    {
        return ConvertStringToDateTime(ReadElementValue(xmlNamespace, localName));
    }

    /// <summary>
    ///     Reads the service objects collection from XML.
    /// </summary>
    /// <typeparam name="TServiceObject">The type of the service object.</typeparam>
    /// <param name="collectionXmlNamespace">Namespace of the collection XML element.</param>
    /// <param name="collectionXmlElementName">Name of the collection XML element.</param>
    /// <param name="getObjectInstanceDelegate">The get object instance delegate.</param>
    /// <param name="clearPropertyBag">if set to <c>true</c> [clear property bag].</param>
    /// <param name="requestedPropertySet">The requested property set.</param>
    /// <param name="summaryPropertiesOnly">if set to <c>true</c> [summary properties only].</param>
    /// <returns>List of service objects.</returns>
    public List<TServiceObject> ReadServiceObjectsCollectionFromXml<TServiceObject>(
        XmlNamespace collectionXmlNamespace,
        string collectionXmlElementName,
        GetObjectInstanceDelegate<TServiceObject?> getObjectInstanceDelegate,
        bool clearPropertyBag,
        PropertySet? requestedPropertySet,
        bool summaryPropertiesOnly
    )
        where TServiceObject : ServiceObject
    {
        var serviceObjects = new List<TServiceObject>();

        if (!IsStartElement(collectionXmlNamespace, collectionXmlElementName))
        {
            ReadStartElement(collectionXmlNamespace, collectionXmlElementName);
        }

        if (!IsEmptyElement)
        {
            do
            {
                Read();

                if (IsStartElement())
                {
                    var serviceObject = getObjectInstanceDelegate(Service, LocalName);

                    if (serviceObject == null)
                    {
                        SkipCurrentElement();
                    }
                    else
                    {
                        if (string.Compare(LocalName, serviceObject.GetXmlElementName(), StringComparison.Ordinal) != 0)
                        {
                            throw new ServiceLocalException(
                                $"The type of the object in the store ({LocalName}) does not match that of the local object ({serviceObject.GetXmlElementName()})."
                            );
                        }

                        serviceObject.LoadFromXml(this, clearPropertyBag, requestedPropertySet, summaryPropertiesOnly);

                        serviceObjects.Add(serviceObject);
                    }
                }
            } while (!IsEndElement(collectionXmlNamespace, collectionXmlElementName));
        }

        return serviceObjects;
    }

    /// <summary>
    ///     Reads the service objects collection from XML.
    /// </summary>
    /// <typeparam name="TServiceObject">The type of the service object.</typeparam>
    /// <param name="collectionXmlElementName">Name of the collection XML element.</param>
    /// <param name="getObjectInstanceDelegate">The get object instance delegate.</param>
    /// <param name="clearPropertyBag">if set to <c>true</c> [clear property bag].</param>
    /// <param name="requestedPropertySet">The requested property set.</param>
    /// <param name="summaryPropertiesOnly">if set to <c>true</c> [summary properties only].</param>
    /// <returns>List of service objects.</returns>
    public List<TServiceObject> ReadServiceObjectsCollectionFromXml<TServiceObject>(
        string collectionXmlElementName,
        GetObjectInstanceDelegate<TServiceObject> getObjectInstanceDelegate,
        bool clearPropertyBag,
        PropertySet? requestedPropertySet,
        bool summaryPropertiesOnly
    )
        where TServiceObject : ServiceObject
    {
        return ReadServiceObjectsCollectionFromXml(
            XmlNamespace.Messages,
            collectionXmlElementName,
            getObjectInstanceDelegate,
            clearPropertyBag,
            requestedPropertySet,
            summaryPropertiesOnly
        );
    }
}
