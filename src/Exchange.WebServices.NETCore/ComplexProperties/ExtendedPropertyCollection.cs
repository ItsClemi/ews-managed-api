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
using System.Diagnostics.CodeAnalysis;
using System.Reflection;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a collection of extended properties.
/// </summary>
[PublicAPI]
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ExtendedPropertyCollection : ComplexPropertyCollection<ExtendedProperty>, ICustomUpdateSerializer
{
    /// <summary>
    ///     Creates the complex property.
    /// </summary>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <returns>Complex property instance.</returns>
    internal override ExtendedProperty CreateComplexProperty(string xmlElementName)
    {
        return new ExtendedProperty();
    }

    /// <summary>
    ///     Gets the name of the collection item XML element.
    /// </summary>
    /// <param name="complexProperty">The complex property.</param>
    /// <returns>XML element name.</returns>
    internal override string GetCollectionItemXmlElementName(ExtendedProperty complexProperty)
    {
        // This method is unused in this class, so just return null.
        return null;
    }

    /// <summary>
    ///     Loads from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="localElementName">Name of the local element.</param>
    internal override void LoadFromXml(EwsServiceXmlReader reader, string localElementName)
    {
        var extendedProperty = new ExtendedProperty();

        extendedProperty.LoadFromXml(reader, reader.LocalName);
        InternalAdd(extendedProperty);
    }

    /// <summary>
    ///     Writes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    internal override void WriteToXml(EwsServiceXmlWriter writer, string xmlElementName)
    {
        foreach (var extendedProperty in this)
        {
            extendedProperty.WriteToXml(writer, XmlElementNames.ExtendedProperty);
        }
    }

    /// <summary>
    ///     Gets existing or adds new extended property.
    /// </summary>
    /// <param name="propertyDefinition">The property definition.</param>
    /// <returns>ExtendedProperty.</returns>
    private ExtendedProperty GetOrAddExtendedProperty(ExtendedPropertyDefinition propertyDefinition)
    {
        if (!TryGetProperty(propertyDefinition, out var extendedProperty))
        {
            extendedProperty = new ExtendedProperty(propertyDefinition);
            InternalAdd(extendedProperty);
        }

        return extendedProperty;
    }

    /// <summary>
    ///     Sets an extended property.
    /// </summary>
    /// <param name="propertyDefinition">The property definition.</param>
    /// <param name="value">The value.</param>
    internal void SetExtendedProperty(ExtendedPropertyDefinition propertyDefinition, object value)
    {
        var extendedProperty = GetOrAddExtendedProperty(propertyDefinition);
        extendedProperty.Value = value;
    }

    /// <summary>
    ///     Removes a specific extended property definition from the collection.
    /// </summary>
    /// <param name="propertyDefinition">The definition of the extended property to remove.</param>
    /// <returns>
    ///     True if the property matching the extended property definition was successfully removed from the collection,
    ///     false otherwise.
    /// </returns>
    internal bool RemoveExtendedProperty(ExtendedPropertyDefinition propertyDefinition)
    {
        EwsUtilities.ValidateParam(propertyDefinition);

        if (TryGetProperty(propertyDefinition, out var extendedProperty))
        {
            return InternalRemove(extendedProperty);
        }

        return false;
    }

    /// <summary>
    ///     Tries to get property.
    /// </summary>
    /// <param name="propertyDefinition">The property definition.</param>
    /// <param name="extendedProperty">The extended property.</param>
    /// <returns>True of property exists in collection.</returns>
    private bool TryGetProperty(
        ExtendedPropertyDefinition propertyDefinition,
        [MaybeNullWhen(false)] out ExtendedProperty extendedProperty
    )
    {
        extendedProperty = Items.Find(prop => prop.PropertyDefinition.Equals(propertyDefinition));
        return extendedProperty != null;
    }

    /// <summary>
    ///     Tries to get property value.
    /// </summary>
    /// <param name="propertyDefinition">The property definition.</param>
    /// <param name="propertyValue">The property value.</param>
    /// <typeparam name="T">Type of expected property value.</typeparam>
    /// <returns>True if property exists in collection.</returns>
    internal bool TryGetValue<T>(
        ExtendedPropertyDefinition propertyDefinition,
        [MaybeNullWhen(false)] out T propertyValue
    )
    {
        if (TryGetProperty(propertyDefinition, out var extendedProperty))
        {
            // Verify that the type parameter and property definition's type are compatible.
            if (!typeof(T).GetTypeInfo().IsAssignableFrom(propertyDefinition.Type.GetTypeInfo()))
            {
                var errorMessage = string.Format(
                    Strings.PropertyDefinitionTypeMismatch,
                    EwsUtilities.GetPrintableTypeName(propertyDefinition.Type),
                    EwsUtilities.GetPrintableTypeName(typeof(T))
                );

                throw new ArgumentException(errorMessage, nameof(propertyDefinition));
            }

            propertyValue = (T)extendedProperty.Value;
            return true;
        }

        propertyValue = default;
        return false;
    }


    #region ICustomXmlUpdateSerializer Members

    /// <summary>
    ///     Writes the update to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="ewsObject">The ews object.</param>
    /// <param name="propertyDefinition">Property definition.</param>
    /// <returns>
    ///     True if property generated serialization.
    /// </returns>
    bool ICustomUpdateSerializer.WriteSetUpdateToXml(
        EwsServiceXmlWriter writer,
        ServiceObject ewsObject,
        PropertyDefinition propertyDefinition
    )
    {
        var propertiesToSet = new List<ExtendedProperty>();

        propertiesToSet.AddRange(AddedItems);
        propertiesToSet.AddRange(ModifiedItems);

        foreach (var extendedProperty in propertiesToSet)
        {
            writer.WriteStartElement(XmlNamespace.Types, ewsObject.GetSetFieldXmlElementName());
            extendedProperty.PropertyDefinition.WriteToXml(writer);

            writer.WriteStartElement(XmlNamespace.Types, ewsObject.GetXmlElementName());
            extendedProperty.WriteToXml(writer, XmlElementNames.ExtendedProperty);
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        foreach (var extendedProperty in RemovedItems)
        {
            writer.WriteStartElement(XmlNamespace.Types, ewsObject.GetDeleteFieldXmlElementName());
            extendedProperty.PropertyDefinition.WriteToXml(writer);
            writer.WriteEndElement();
        }

        return true;
    }

    /// <summary>
    ///     Writes the deletion update to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="ewsObject">The ews object.</param>
    /// <returns>
    ///     True if property generated serialization.
    /// </returns>
    bool ICustomUpdateSerializer.WriteDeleteUpdateToXml(EwsServiceXmlWriter writer, ServiceObject ewsObject)
    {
        foreach (var extendedProperty in Items)
        {
            writer.WriteStartElement(XmlNamespace.Types, ewsObject.GetDeleteFieldXmlElementName());
            extendedProperty.PropertyDefinition.WriteToXml(writer);
            writer.WriteEndElement();
        }

        return true;
    }

    #endregion
}
