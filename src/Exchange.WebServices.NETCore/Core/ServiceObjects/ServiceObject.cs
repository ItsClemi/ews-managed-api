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
using System.Diagnostics.CodeAnalysis;
using System.Reflection;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents the base abstract class for all item and folder types.
/// </summary>
[PublicAPI]
public abstract class ServiceObject
{
    private readonly object _lockObject = new();
    private string _xmlElementName;

    /// <summary>
    ///     The property bag holding property values for this object.
    /// </summary>
    internal PropertyBag PropertyBag { get; }

    /// <summary>
    ///     Gets the schema associated with this type of object.
    /// </summary>
    public ServiceObjectSchema Schema => GetSchema();

    /// <summary>
    ///     Gets the value of specified property in this instance.
    /// </summary>
    /// <param name="propertyDefinition">Definition of the property to get.</param>
    /// <exception cref="ServiceVersionException">Raised if this property requires a later version of Exchange.</exception>
    /// <exception cref="PropertyException">
    ///     Raised if this property hasn't been assigned or loaded. Raised for set if property
    ///     cannot be updated or deleted.
    /// </exception>
    public object? this[PropertyDefinitionBase propertyDefinition]
    {
        get
        {
            if (propertyDefinition is PropertyDefinition propDef)
            {
                return PropertyBag[propDef];
            }

            var extendedPropDef = propertyDefinition as ExtendedPropertyDefinition;
            if (extendedPropDef != null)
            {
                if (TryGetExtendedProperty(extendedPropDef, out object? propertyValue))
                {
                    return propertyValue;
                }

                throw new ServiceObjectPropertyException(
                    Strings.MustLoadOrAssignPropertyBeforeAccess,
                    propertyDefinition
                );
            }

            // Other subclasses of PropertyDefinitionBase are not supported.
            throw new NotSupportedException(
                string.Format(Strings.OperationNotSupportedForPropertyDefinitionType, propertyDefinition.GetType().Name)
            );
        }
    }

    /// <summary>
    ///     Gets the ExchangeService the object is bound to.
    /// </summary>
    public ExchangeService Service { get; internal set; }

    /// <summary>
    ///     Indicates whether this object is a real store item, or if it's a local object
    ///     that has yet to be saved.
    /// </summary>
    public virtual bool IsNew
    {
        get
        {
            var id = GetId();

            return id == null ? true : !id.IsValid;
        }
    }

    /// <summary>
    ///     Gets a value indicating whether the object has been modified and should be saved.
    /// </summary>
    public bool IsDirty => PropertyBag.IsDirty;

    /// <summary>
    ///     Internal constructor.
    /// </summary>
    /// <param name="service">EWS service to which this object belongs.</param>
    internal ServiceObject(ExchangeService service)
    {
        EwsUtilities.ValidateParam(service);
        EwsUtilities.ValidateServiceObjectVersion(this, service.RequestedServerVersion);

        Service = service;
        PropertyBag = new PropertyBag(this);
    }

    /// <summary>
    ///     Triggers dispatch of the change event.
    /// </summary>
    internal void Changed()
    {
        OnChange?.Invoke(this);
    }

    /// <summary>
    ///     Throws exception if this is a new service object.
    /// </summary>
    internal void ThrowIfThisIsNew()
    {
        if (IsNew)
        {
            throw new InvalidOperationException(Strings.ServiceObjectDoesNotHaveId);
        }
    }

    /// <summary>
    ///     Throws exception if this is not a new service object.
    /// </summary>
    internal void ThrowIfThisIsNotNew()
    {
        if (!IsNew)
        {
            throw new InvalidOperationException(Strings.ServiceObjectAlreadyHasId);
        }
    }

    /// <summary>
    ///     This methods lets subclasses of ServiceObject override the default mechanism
    ///     by which the XML element name associated with their type is retrieved.
    /// </summary>
    /// <returns>
    ///     The XML element name associated with this type.
    ///     If this method returns null or empty, the XML element name associated with this
    ///     type is determined by the EwsObjectDefinition attribute that decorates the type,
    ///     if present.
    /// </returns>
    /// <remarks>
    ///     Item and folder classes that can be returned by EWS MUST rely on the EwsObjectDefinition
    ///     attribute for XML element name determination.
    /// </remarks>
    internal virtual string GetXmlElementNameOverride()
    {
        return null!;
    }

    /// <summary>
    ///     GetXmlElementName retrieves the XmlElementName of this type based on the
    ///     EwsObjectDefinition attribute that decorates it, if present.
    /// </summary>
    /// <returns>The XML element name associated with this type.</returns>
    internal string GetXmlElementName()
    {
        if (string.IsNullOrEmpty(_xmlElementName))
        {
            _xmlElementName = GetXmlElementNameOverride();

            if (string.IsNullOrEmpty(_xmlElementName))
            {
                lock (_lockObject)
                {
                    foreach (Attribute attribute in GetType().GetTypeInfo().GetCustomAttributes(false))
                    {
                        if (attribute is ServiceObjectDefinitionAttribute definitionAttribute)
                        {
                            _xmlElementName = definitionAttribute.XmlElementName;
                        }
                    }
                }
            }
        }

        EwsUtilities.Assert(
            !string.IsNullOrEmpty(_xmlElementName),
            "EwsObject.GetXmlElementName",
            string.Format("The class {0} does not have an associated XML element name.", GetType().Name)
        );

        return _xmlElementName;
    }

    /// <summary>
    ///     Gets the name of the change XML element.
    /// </summary>
    /// <returns>XML element name,</returns>
    internal virtual string GetChangeXmlElementName()
    {
        return XmlElementNames.ItemChange;
    }

    /// <summary>
    ///     Gets the name of the set field XML element.
    /// </summary>
    /// <returns>XML element name,</returns>
    internal virtual string GetSetFieldXmlElementName()
    {
        return XmlElementNames.SetItemField;
    }

    /// <summary>
    ///     Gets the name of the delete field XML element.
    /// </summary>
    /// <returns>XML element name,</returns>
    internal virtual string GetDeleteFieldXmlElementName()
    {
        return XmlElementNames.DeleteItemField;
    }

    /// <summary>
    ///     Gets a value indicating whether a time zone SOAP header should be emitted in a CreateItem
    ///     or UpdateItem request so this item can be property saved or updated.
    /// </summary>
    /// <param name="isUpdateOperation">Indicates whether the operation being performed is an update operation.</param>
    /// <returns><c>true</c> if a time zone SOAP header should be emitted; otherwise, <c>false</c>.</returns>
    internal virtual bool GetIsTimeZoneHeaderRequired(bool isUpdateOperation)
    {
        return false;
    }

    /// <summary>
    ///     Determines whether properties defined with ScopedDateTimePropertyDefinition require custom time zone scoping.
    /// </summary>
    /// <returns>
    ///     <c>true</c> if this item type requires custom scoping for scoped date/time properties; otherwise, <c>false</c>.
    /// </returns>
    internal virtual bool GetIsCustomDateTimeScopingRequired()
    {
        return false;
    }

    /// <summary>
    ///     Internal method to return the schema associated with this type of object.
    /// </summary>
    /// <returns>The schema associated with this type of object.</returns>
    internal abstract ServiceObjectSchema GetSchema();

    /// <summary>
    ///     Gets the minimum required server version.
    /// </summary>
    /// <returns>Earliest Exchange version in which this service object type is supported.</returns>
    internal abstract ExchangeVersion GetMinimumRequiredServerVersion();

    /// <summary>
    ///     Loads service object from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="clearPropertyBag">if set to <c>true</c> [clear property bag].</param>
    internal void LoadFromXml(EwsServiceXmlReader reader, bool clearPropertyBag)
    {
        PropertyBag.LoadFromXml(reader, clearPropertyBag, null, false);
    }

    /// <summary>
    ///     Validates this instance.
    /// </summary>
    internal virtual void Validate()
    {
        PropertyBag.Validate();
    }

    /// <summary>
    ///     Loads service object from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="clearPropertyBag">if set to <c>true</c> [clear property bag].</param>
    /// <param name="requestedPropertySet">The property set.</param>
    /// <param name="summaryPropertiesOnly">if set to <c>true</c> [summary props only].</param>
    internal void LoadFromXml(
        EwsServiceXmlReader reader,
        bool clearPropertyBag,
        PropertySet? requestedPropertySet,
        bool summaryPropertiesOnly
    )
    {
        PropertyBag.LoadFromXml(reader, clearPropertyBag, requestedPropertySet, summaryPropertiesOnly);
    }

    /// <summary>
    ///     Clears the object's change log.
    /// </summary>
    internal void ClearChangeLog()
    {
        PropertyBag.ClearChangeLog();
    }

    /// <summary>
    ///     Writes service object as XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal void WriteToXml(EwsServiceXmlWriter writer)
    {
        PropertyBag.WriteToXml(writer);
    }

    /// <summary>
    ///     Writes service object for update as XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal void WriteToXmlForUpdate(EwsServiceXmlWriter writer)
    {
        PropertyBag.WriteToXmlForUpdate(writer);
    }

    /// <summary>
    ///     Loads the specified set of properties on the object.
    /// </summary>
    /// <param name="propertySet">The properties to load.</param>
    /// <param name="token"></param>
    internal abstract Task<ServiceResponseCollection<ServiceResponse>> InternalLoad(
        PropertySet propertySet,
        CancellationToken token
    );

    /// <summary>
    ///     Deletes the object.
    /// </summary>
    /// <param name="deleteMode">The deletion mode.</param>
    /// <param name="sendCancellationsMode">Indicates whether meeting cancellation messages should be sent.</param>
    /// <param name="affectedTaskOccurrences">Indicate which occurrence of a recurring task should be deleted.</param>
    /// <param name="token"></param>
    internal abstract Task<ServiceResponseCollection<ServiceResponse>> InternalDelete(
        DeleteMode deleteMode,
        SendCancellationsMode? sendCancellationsMode,
        AffectedTaskOccurrence? affectedTaskOccurrences,
        CancellationToken token
    );

    /// <summary>
    ///     Loads the specified set of properties. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="propertySet">The properties to load.</param>
    /// <param name="token"></param>
    public Task<ServiceResponseCollection<ServiceResponse>> Load(
        PropertySet propertySet,
        CancellationToken token = default
    )
    {
        return InternalLoad(propertySet, token);
    }

    /// <summary>
    ///     Loads the first class properties. Calling this method results in a call to EWS.
    /// </summary>
    public Task<ServiceResponseCollection<ServiceResponse>> Load(CancellationToken token = default)
    {
        return InternalLoad(PropertySet.FirstClassProperties, token);
    }

    /// <summary>
    ///     Try to get the value of a specified extended property in this instance.
    /// </summary>
    /// <param name="propertyDefinition">The property definition.</param>
    /// <param name="propertyValue">The property value.</param>
    /// <typeparam name="T">Type of expected property value.</typeparam>
    /// <returns>True if property retrieved, false otherwise.</returns>
    internal bool TryGetExtendedProperty<T>(
        ExtendedPropertyDefinition propertyDefinition,
        [MaybeNullWhen(false)] out T propertyValue
    )
    {
        var propertyCollection = GetExtendedProperties();

        if (propertyCollection != null && propertyCollection.TryGetValue(propertyDefinition, out propertyValue))
        {
            return true;
        }

        propertyValue = default;
        return false;
    }

    /// <summary>
    ///     Try to get the value of a specified property in this instance.
    /// </summary>
    /// <param name="propertyDefinition">The property definition.</param>
    /// <param name="propertyValue">The property value.</param>
    /// <returns>True if property retrieved, false otherwise.</returns>
    public bool TryGetProperty(
        PropertyDefinitionBase propertyDefinition,
        [MaybeNullWhen(false)] out object propertyValue
    )
    {
        return TryGetProperty<object>(propertyDefinition, out propertyValue);
    }

    /// <summary>
    ///     Try to get the value of a specified property in this instance.
    /// </summary>
    /// <param name="propertyDefinition">The property definition.</param>
    /// <param name="propertyValue">The property value.</param>
    /// <typeparam name="T">Type of expected property value.</typeparam>
    /// <returns>True if property retrieved, false otherwise.</returns>
    public bool TryGetProperty<T>(PropertyDefinitionBase propertyDefinition, [MaybeNullWhen(false)] out T propertyValue)
    {
        if (propertyDefinition is PropertyDefinition propDef)
        {
            return PropertyBag.TryGetProperty(propDef, out propertyValue);
        }

        var extPropDef = propertyDefinition as ExtendedPropertyDefinition;
        if (extPropDef != null)
        {
            return TryGetExtendedProperty(extPropDef, out propertyValue);
        }

        // Other subclasses of PropertyDefinitionBase are not supported.
        throw new NotSupportedException(
            string.Format(Strings.OperationNotSupportedForPropertyDefinitionType, propertyDefinition.GetType().Name)
        );
    }

    /// <summary>
    ///     Gets the collection of loaded property definitions.
    /// </summary>
    /// <returns>Collection of property definitions.</returns>
    public Collection<PropertyDefinitionBase> GetLoadedPropertyDefinitions()
    {
        var propDefs = new Collection<PropertyDefinitionBase>();
        foreach (var propDef in PropertyBag.Properties.Keys)
        {
            propDefs.Add(propDef);
        }

        var properties = GetExtendedProperties();
        if (properties != null)
        {
            foreach (var extProp in properties)
            {
                propDefs.Add(extProp.PropertyDefinition);
            }
        }

        return propDefs;
    }

    /// <summary>
    ///     The property definition for the Id of this object.
    /// </summary>
    /// <returns>A PropertyDefinition instance.</returns>
    internal virtual PropertyDefinition? GetIdPropertyDefinition()
    {
        return null;
    }

    /// <summary>
    ///     The unique Id of this object.
    /// </summary>
    /// <returns>A ServiceId instance.</returns>
    internal ServiceId GetId()
    {
        var idPropertyDefinition = GetIdPropertyDefinition();

        object? serviceId = null;

        if (idPropertyDefinition != null)
        {
            PropertyBag.TryGetValue(idPropertyDefinition, out serviceId);
        }

        return (ServiceId)serviceId;
    }

    /// <summary>
    ///     Gets the extended properties collection.
    /// </summary>
    /// <returns>Extended properties collection.</returns>
    internal virtual ExtendedPropertyCollection? GetExtendedProperties()
    {
        return null;
    }

    /// <summary>
    ///     Defines an event that is triggered when the service object changes.
    /// </summary>
    internal event ServiceObjectChangedDelegate? OnChange;
}
