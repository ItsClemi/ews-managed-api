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
///     Represents an entry of an PhysicalAddressDictionary.
/// </summary>
[PublicAPI]
public sealed class PhysicalAddressEntry : DictionaryEntryProperty<PhysicalAddressKey>
{
    private readonly SimplePropertyBag<string> _propertyBag;


    /// <summary>
    ///     Initializes a new instance of PhysicalAddressEntry
    /// </summary>
    public PhysicalAddressEntry()
    {
        _propertyBag = new SimplePropertyBag<string>();
        _propertyBag.OnChange += PropertyBagChanged;
    }


    #region Classes

    /// <summary>
    ///     Schema definition for PhysicalAddress
    /// </summary>
    private static class PhysicalAddressSchema
    {
        public const string Street = "Street";
        public const string City = "City";
        public const string State = "State";
        public const string CountryOrRegion = "CountryOrRegion";
        public const string PostalCode = "PostalCode";


        /// <summary>
        ///     Gets the XML element names.
        /// </summary>
        /// <value>The XML element names.</value>
        public static List<string> XmlElementNames =>
            new List<string>
            {
                Street,
                City,
                State,
                CountryOrRegion,
                PostalCode,
            };
    }

    #endregion


    #region Properties

    /// <summary>
    ///     Gets or sets the street.
    /// </summary>
    public string Street
    {
        get => (string)_propertyBag[PhysicalAddressSchema.Street];
        set => _propertyBag[PhysicalAddressSchema.Street] = value;
    }

    /// <summary>
    ///     Gets or sets the city.
    /// </summary>
    public string City
    {
        get => (string)_propertyBag[PhysicalAddressSchema.City];
        set => _propertyBag[PhysicalAddressSchema.City] = value;
    }

    /// <summary>
    ///     Gets or sets the state.
    /// </summary>
    public string State
    {
        get => (string)_propertyBag[PhysicalAddressSchema.State];
        set => _propertyBag[PhysicalAddressSchema.State] = value;
    }

    /// <summary>
    ///     Gets or sets the country or region.
    /// </summary>
    public string CountryOrRegion
    {
        get => (string)_propertyBag[PhysicalAddressSchema.CountryOrRegion];
        set => _propertyBag[PhysicalAddressSchema.CountryOrRegion] = value;
    }

    /// <summary>
    ///     Gets or sets the postal code.
    /// </summary>
    public string PostalCode
    {
        get => (string)_propertyBag[PhysicalAddressSchema.PostalCode];
        set => _propertyBag[PhysicalAddressSchema.PostalCode] = value;
    }

    #endregion


    #region Internal methods

    /// <summary>
    ///     Clears the change log.
    /// </summary>
    internal override void ClearChangeLog()
    {
        _propertyBag.ClearChangeLog();
    }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        if (PhysicalAddressSchema.XmlElementNames.Contains(reader.LocalName))
        {
            _propertyBag[reader.LocalName] = reader.ReadElementValue();

            return true;
        }

        return false;
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        foreach (var xmlElementName in PhysicalAddressSchema.XmlElementNames)
        {
            writer.WriteElementValue(XmlNamespace.Types, xmlElementName, _propertyBag[xmlElementName]);
        }
    }

    /// <summary>
    ///     Writes the update to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="ewsObject">The ews object.</param>
    /// <param name="ownerDictionaryXmlElementName">Name of the owner dictionary XML element.</param>
    /// <returns>True if update XML was written.</returns>
    internal override bool WriteSetUpdateToXml(
        EwsServiceXmlWriter writer,
        ServiceObject ewsObject,
        string ownerDictionaryXmlElementName
    )
    {
        var fieldsToSet = new List<string>();

        foreach (var xmlElementName in _propertyBag.AddedItems)
        {
            fieldsToSet.Add(xmlElementName);
        }

        foreach (var xmlElementName in _propertyBag.ModifiedItems)
        {
            fieldsToSet.Add(xmlElementName);
        }

        foreach (var xmlElementName in fieldsToSet)
        {
            writer.WriteStartElement(XmlNamespace.Types, ewsObject.GetSetFieldXmlElementName());

            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.IndexedFieldURI);
            writer.WriteAttributeValue(XmlAttributeNames.FieldURI, GetFieldUri(xmlElementName));
            writer.WriteAttributeValue(XmlAttributeNames.FieldIndex, Key.ToString());
            writer.WriteEndElement(); // IndexedFieldURI

            writer.WriteStartElement(XmlNamespace.Types, ewsObject.GetXmlElementName());
            writer.WriteStartElement(XmlNamespace.Types, ownerDictionaryXmlElementName);
            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Entry);
            WriteAttributesToXml(writer);
            writer.WriteElementValue(XmlNamespace.Types, xmlElementName, _propertyBag[xmlElementName]);
            writer.WriteEndElement(); // Entry
            writer.WriteEndElement(); // ownerDictionaryXmlElementName
            writer.WriteEndElement(); // ewsObject.GetXmlElementName()

            writer.WriteEndElement(); // ewsObject.GetSetFieldXmlElementName()
        }

        foreach (var xmlElementName in _propertyBag.RemovedItems)
        {
            InternalWriteDeleteFieldToXml(writer, ewsObject, xmlElementName);
        }

        return true;
    }

    /// <summary>
    ///     Writes the delete update to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="ewsObject">The ews object.</param>
    /// <returns>True if update XML was written.</returns>
    internal override bool WriteDeleteUpdateToXml(EwsServiceXmlWriter writer, ServiceObject ewsObject)
    {
        foreach (var xmlElementName in PhysicalAddressSchema.XmlElementNames)
        {
            InternalWriteDeleteFieldToXml(writer, ewsObject, xmlElementName);
        }

        return true;
    }

    #endregion


    #region Private methods

    /// <summary>
    ///     Gets the field URI.
    /// </summary>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <returns>Field URI.</returns>
    private static string GetFieldUri(string xmlElementName)
    {
        return "contacts:PhysicalAddress:" + xmlElementName;
    }

    /// <summary>
    ///     Property bag was changed.
    /// </summary>
    private void PropertyBagChanged()
    {
        Changed();
    }

    /// <summary>
    ///     Write field deletion to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="ewsObject">The ews object.</param>
    /// <param name="fieldXmlElementName">Name of the field XML element.</param>
    private void InternalWriteDeleteFieldToXml(
        EwsServiceXmlWriter writer,
        ServiceObject ewsObject,
        string fieldXmlElementName
    )
    {
        writer.WriteStartElement(XmlNamespace.Types, ewsObject.GetDeleteFieldXmlElementName());
        writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.IndexedFieldURI);
        writer.WriteAttributeValue(XmlAttributeNames.FieldURI, GetFieldUri(fieldXmlElementName));
        writer.WriteAttributeValue(XmlAttributeNames.FieldIndex, Key.ToString());
        writer.WriteEndElement(); // IndexedFieldURI
        writer.WriteEndElement(); // ewsObject.GetDeleteFieldXmlElementName()
    }

    #endregion
}
