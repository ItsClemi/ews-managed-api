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
///     Represents the definition of a folder or item property.
/// </summary>
[PublicAPI]
public abstract class PropertyDefinition : ServiceObjectPropertyDefinition
{
    private readonly PropertyDefinitionFlags _flags;
    private string _name;

    /// <summary>
    ///     Gets the minimum Exchange version that supports this property.
    /// </summary>
    /// <value>The version.</value>
    public override ExchangeVersion Version { get; }

    /// <summary>
    ///     Gets a value indicating whether this property definition is for a nullable type (ref, int?, bool?...).
    /// </summary>
    internal virtual bool IsNullable => true;

    /// <summary>
    ///     Gets the name of the XML element.
    /// </summary>
    /// <value>The name of the XML element.</value>
    internal string XmlElementName { get; }

    /// <summary>
    ///     Gets the name of the property.
    /// </summary>
    public string Name
    {
        get
        {
            // Name is initialized at read time for all PropertyDefinition instances using Reflection.
            if (string.IsNullOrEmpty(_name))
            {
                ServiceObjectSchema.InitializeSchemaPropertyNames();
            }

            return _name;
        }

        internal set => _name = value;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="PropertyDefinition" /> class.
    /// </summary>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <param name="uri">The URI.</param>
    /// <param name="version">The version.</param>
    internal PropertyDefinition(string xmlElementName, string uri, ExchangeVersion version)
        : base(uri)
    {
        XmlElementName = xmlElementName;
        _flags = PropertyDefinitionFlags.None;
        Version = version;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="PropertyDefinition" /> class.
    /// </summary>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <param name="flags">The flags.</param>
    /// <param name="version">The version.</param>
    internal PropertyDefinition(string xmlElementName, PropertyDefinitionFlags flags, ExchangeVersion version)
    {
        XmlElementName = xmlElementName;
        _flags = flags;
        Version = version;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="PropertyDefinition" /> class.
    /// </summary>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <param name="uri">The URI.</param>
    /// <param name="flags">The flags.</param>
    /// <param name="version">The version.</param>
    internal PropertyDefinition(
        string xmlElementName,
        string uri,
        PropertyDefinitionFlags flags,
        ExchangeVersion version
    )
        : this(xmlElementName, uri, version)
    {
        _flags = flags;
    }

    /// <summary>
    ///     Determines whether the specified flag is set.
    /// </summary>
    /// <param name="flag">The flag.</param>
    /// <returns>
    ///     <c>true</c> if the specified flag is set; otherwise, <c>false</c>.
    /// </returns>
    internal bool HasFlag(PropertyDefinitionFlags flag)
    {
        return HasFlag(flag, null);
    }

    /// <summary>
    ///     Determines whether the specified flag is set.
    /// </summary>
    /// <param name="flag">The flag.</param>
    /// <param name="version">Requested version.</param>
    /// <returns>
    ///     <c>true</c> if the specified flag is set; otherwise, <c>false</c>.
    /// </returns>
    internal virtual bool HasFlag(PropertyDefinitionFlags flag, ExchangeVersion? version)
    {
        return (_flags & flag) == flag;
    }

    /// <summary>
    ///     Registers associated internal properties.
    /// </summary>
    /// <param name="properties">The list in which to add the associated properties.</param>
    internal virtual void RegisterAssociatedInternalProperties(List<PropertyDefinition> properties)
    {
    }

    /// <summary>
    ///     Gets a list of associated internal properties.
    /// </summary>
    /// <returns>A list of PropertyDefinition objects.</returns>
    /// <remarks>
    ///     This is a hack. It is here (currently) solely to help the API
    ///     register the MeetingTimeZone property definition that is internal.
    /// </remarks>
    internal List<PropertyDefinition> GetAssociatedInternalProperties()
    {
        var properties = new List<PropertyDefinition>();

        RegisterAssociatedInternalProperties(properties);

        return properties;
    }

    /// <summary>
    ///     Loads from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="propertyBag">The property bag.</param>
    internal abstract void LoadPropertyValueFromXml(EwsServiceXmlReader reader, PropertyBag propertyBag);

    /// <summary>
    ///     Writes the property value to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="propertyBag">The property bag.</param>
    /// <param name="isUpdateOperation">Indicates whether the context is an update operation.</param>
    internal abstract void WritePropertyValueToXml(
        EwsServiceXmlWriter writer,
        PropertyBag propertyBag,
        bool isUpdateOperation
    );

    /// <summary>
    ///     Gets the property definition's printable name.
    /// </summary>
    /// <returns>
    ///     The property definition's printable name.
    /// </returns>
    internal override string GetPrintableName()
    {
        return Name;
    }
}
