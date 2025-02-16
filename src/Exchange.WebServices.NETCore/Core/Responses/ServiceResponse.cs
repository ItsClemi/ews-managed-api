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

using System.Collections.ObjectModel;
using System.Net.Http.Headers;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents the standard response to an Exchange Web Services operation.
/// </summary>
[PublicAPI]
public class ServiceResponse
{
    private readonly Dictionary<string, string> _errorDetails = [];

    /// <summary>
    ///     Gets a value indicating whether a batch request stopped processing before the end.
    /// </summary>
    internal bool BatchProcessingStopped =>
        Result == ServiceResult.Warning && ErrorCode == ServiceError.ErrorBatchProcessingStopped;

    /// <summary>
    ///     Gets the result associated with this response.
    /// </summary>
    public ServiceResult Result { get; private set; }

    /// <summary>
    ///     Gets the error code associated with this response.
    /// </summary>
    public ServiceError ErrorCode { get; private set; }

    /// <summary>
    ///     Gets a detailed error message associated with the response. If Result is set to Success, ErrorMessage returns null.
    ///     ErrorMessage is localized according to the PreferredCulture property of the ExchangeService object that
    ///     was used to call the method that generated the response.
    /// </summary>
    public string ErrorMessage { get; internal set; }

    /// <summary>
    ///     Gets error details associated with the response. If Result is set to Success, ErrorDetailsDictionary returns null.
    ///     Error details will only available for some error codes. For example, when error code is
    ///     ErrorRecurrenceHasNoOccurrence,
    ///     the ErrorDetailsDictionary will contain keys for EffectiveStartDate and EffectiveEndDate.
    /// </summary>
    /// <value>The error details dictionary.</value>
    public IDictionary<string, string> ErrorDetails => _errorDetails;

    /// <summary>
    ///     Gets information about property errors associated with the response. If Result is set to Success, ErrorProperties
    ///     returns null.
    ///     ErrorProperties is only available for some error codes. For example, when the error code is
    ///     ErrorInvalidPropertyForOperation,
    ///     ErrorProperties will contain the definition of the property that was invalid for the request.
    /// </summary>
    /// <value>The error properties list.</value>
    public Collection<PropertyDefinitionBase> ErrorProperties { get; } = new();

    /// <summary>
    ///     Initializes a new instance of the <see cref="ServiceResponse" /> class.
    /// </summary>
    internal ServiceResponse()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ServiceResponse" /> class.
    /// </summary>
    /// <param name="soapFaultDetails">The SOAP fault details.</param>
    internal ServiceResponse(SoapFaultDetails soapFaultDetails)
    {
        Result = ServiceResult.Error;
        ErrorCode = soapFaultDetails.ResponseCode;
        ErrorMessage = soapFaultDetails.FaultString;
        _errorDetails = soapFaultDetails.ErrorDetails;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ServiceResponse" /> class.
    ///     This is intended to be used by unit tests to create a fake service error response
    /// </summary>
    /// <param name="responseCode">Response code</param>
    /// <param name="errorMessage">Detailed error message</param>
    internal ServiceResponse(ServiceError responseCode, string errorMessage)
    {
        Result = ServiceResult.Error;
        ErrorCode = responseCode;
        ErrorMessage = errorMessage;
        _errorDetails = null;
    }

    /// <summary>
    ///     Loads response from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    internal void LoadFromXml(EwsServiceXmlReader reader, string xmlElementName)
    {
        if (!reader.IsStartElement(XmlNamespace.Messages, xmlElementName))
        {
            reader.ReadStartElement(XmlNamespace.Messages, xmlElementName);
        }

        Result = reader.ReadAttributeValue<ServiceResult>(XmlAttributeNames.ResponseClass);

        if (Result == ServiceResult.Success || Result == ServiceResult.Warning)
        {
            if (Result == ServiceResult.Warning)
            {
                ErrorMessage = reader.ReadElementValue(XmlNamespace.Messages, XmlElementNames.MessageText);
            }

            ErrorCode = reader.ReadElementValue<ServiceError>(XmlNamespace.Messages, XmlElementNames.ResponseCode);

            if (Result == ServiceResult.Warning)
            {
                reader.ReadElementValue<int>(XmlNamespace.Messages, XmlElementNames.DescriptiveLinkKey);
            }
            // TODO: is this required from: https://github.com/sherlock1982/ews-managed-api/pull/42/files
            //else
            //{
            //    reader.ReadElementValue();
            //    if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.MessageText))
            //    {
            //        ErrorMessage = reader.ReadElementValue(XmlNamespace.Messages, XmlElementNames.MessageText);
            //    }
            //}

            // If batch processing stopped, EWS returns an empty element. Skip over it.
            if (BatchProcessingStopped)
            {
                do
                {
                    reader.Read();
                } while (!reader.IsEndElement(XmlNamespace.Messages, xmlElementName));
            }
            else
            {
                ReadElementsFromXml(reader);

                reader.ReadEndElementIfNecessary(XmlNamespace.Messages, xmlElementName);
            }
        }
        else
        {
            ErrorMessage = reader.ReadElementValue(XmlNamespace.Messages, XmlElementNames.MessageText);
            ErrorCode = reader.ReadElementValue<ServiceError>(XmlNamespace.Messages, XmlElementNames.ResponseCode);
            reader.ReadElementValue<int>(XmlNamespace.Messages, XmlElementNames.DescriptiveLinkKey);

            while (!reader.IsEndElement(XmlNamespace.Messages, xmlElementName))
            {
                reader.Read();

                if (reader.IsStartElement())
                {
                    if (!LoadExtraErrorDetailsFromXml(reader, reader.LocalName))
                    {
                        reader.SkipCurrentElement();
                    }
                }
            }
        }

        MapErrorCodeToErrorMessage();

        Loaded();
    }

    /// <summary>
    ///     Parses the message XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    private void ParseMessageXml(EwsServiceXmlReader reader)
    {
        do
        {
            reader.Read();

            if (reader.IsStartElement())
            {
                switch (reader.LocalName)
                {
                    case XmlElementNames.Value:
                    {
                        _errorDetails.Add(reader.ReadAttributeValue(XmlAttributeNames.Name), reader.ReadElementValue());
                        break;
                    }
                    case XmlElementNames.FieldURI:
                    {
                        ErrorProperties.Add(
                            ServiceObjectSchema.FindPropertyDefinition(
                                reader.ReadAttributeValue(XmlAttributeNames.FieldURI)
                            )
                        );
                        break;
                    }
                    case XmlElementNames.IndexedFieldURI:
                    {
                        ErrorProperties.Add(
                            new IndexedPropertyDefinition(
                                reader.ReadAttributeValue(XmlAttributeNames.FieldURI),
                                reader.ReadAttributeValue(XmlAttributeNames.FieldIndex)
                            )
                        );
                        break;
                    }
                    case XmlElementNames.ExtendedFieldURI:
                    {
                        var extendedPropDef = new ExtendedPropertyDefinition();
                        extendedPropDef.LoadFromXml(reader);
                        ErrorProperties.Add(extendedPropDef);
                        break;
                    }
                }
            }
        } while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.MessageXml));
    }

    /// <summary>
    ///     Called when the response has been loaded from XML.
    /// </summary>
    internal virtual void Loaded()
    {
    }

    /// <summary>
    ///     Called after the response has been loaded from XML in order to map error codes to "better" error messages.
    /// </summary>
    internal void MapErrorCodeToErrorMessage()
    {
        // Use a better error message when an item cannot be updated because its changeKey is old.
        if (ErrorCode == ServiceError.ErrorIrresolvableConflict)
        {
            ErrorMessage = Strings.ItemIsOutOfDate;
        }
    }

    /// <summary>
    ///     Reads response elements from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal virtual void ReadElementsFromXml(EwsServiceXmlReader reader)
    {
    }

    /// <summary>
    ///     Reads the headers from a HTTP response
    /// </summary>
    /// <param name="responseHeaders">a collection of response headers</param>
    internal virtual void ReadHeader(HttpResponseHeaders responseHeaders)
    {
    }

    /// <summary>
    ///     Loads extra error details from XML
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="xmlElementName">The current element name of the extra error details.</param>
    /// <returns>
    ///     True if the expected extra details is loaded;
    ///     False if the element name does not match the expected element.
    /// </returns>
    internal virtual bool LoadExtraErrorDetailsFromXml(EwsServiceXmlReader reader, string xmlElementName)
    {
        if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.MessageXml) && !reader.IsEmptyElement)
        {
            ParseMessageXml(reader);

            return true;
        }

        return false;
    }

    /// <summary>
    ///     Throws a ServiceResponseException if this response has its Result property set to Error.
    /// </summary>
    internal void ThrowIfNecessary()
    {
        InternalThrowIfNecessary();
    }

    /// <summary>
    ///     Internal method that throws a ServiceResponseException if this response has its Result property set to Error.
    /// </summary>
    internal virtual void InternalThrowIfNecessary()
    {
        if (Result == ServiceResult.Error)
        {
            throw new ServiceResponseException(this);
        }
    }
}
