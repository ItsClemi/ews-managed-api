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

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents an object that can be used to store user-defined configuration settings.
/// </summary>
[PublicAPI]
public class UserConfiguration
{
    private const ExchangeVersion ObjectVersion = ExchangeVersion.Exchange2010;

    // For consistency with ServiceObject behavior, access to ItemId is permitted for a new object.
    private const UserConfigurationProperties PropertiesAvailableForNewObject = UserConfigurationProperties.BinaryData |
        UserConfigurationProperties.Dictionary |
        UserConfigurationProperties.XmlData;

    private const UserConfigurationProperties NoProperties = 0;

    // TODO: Consider using SimplePropertyBag class to store XmlData & BinaryData property values.
    private readonly ExchangeService _service;
    private byte[]? _binaryData;

    /// <summary>
    ///     Indicates whether changes trigger an update or create operation.
    /// </summary>
    private bool _isNew;

    private UserConfigurationProperties _propertiesAvailableForAccess;
    private UserConfigurationProperties _updatedProperties;
    private byte[]? _xmlData;

    /// <summary>
    ///     Gets the name of the user configuration.
    /// </summary>
    public string Name { get; internal set; }

    /// <summary>
    ///     Gets the Id of the folder containing the user configuration.
    /// </summary>
    public FolderId ParentFolderId { get; internal set; }

    /// <summary>
    ///     Gets the Id of the user configuration.
    /// </summary>
    public ItemId? ItemId { get; private set; }

    /// <summary>
    ///     Gets the dictionary of the user configuration.
    /// </summary>
    public UserConfigurationDictionary Dictionary { get; private set; }

    /// <summary>
    ///     Gets or sets the xml data of the user configuration.
    /// </summary>
    public byte[]? XmlData
    {
        get
        {
            ValidatePropertyAccess(UserConfigurationProperties.XmlData);
            return _xmlData;
        }

        set
        {
            _xmlData = value;
            MarkPropertyForUpdate(UserConfigurationProperties.XmlData);
        }
    }

    /// <summary>
    ///     Gets or sets the binary data of the user configuration.
    /// </summary>
    public byte[]? BinaryData
    {
        get
        {
            ValidatePropertyAccess(UserConfigurationProperties.BinaryData);
            return _binaryData;
        }

        set
        {
            _binaryData = value;
            MarkPropertyForUpdate(UserConfigurationProperties.BinaryData);
        }
    }

    /// <summary>
    ///     Gets a value indicating whether this user configuration has been modified.
    /// </summary>
    public bool IsDirty => _updatedProperties != NoProperties || Dictionary.IsDirty;

    /// <summary>
    ///     Initializes a new instance of <see cref="UserConfiguration" /> class.
    /// </summary>
    /// <param name="service">The service to which the user configuration is bound.</param>
    public UserConfiguration(ExchangeService service)
        : this(service, PropertiesAvailableForNewObject)
    {
    }

    /// <summary>
    ///     Initializes a new instance of <see cref="UserConfiguration" /> class.
    /// </summary>
    /// <param name="service">The service to which the user configuration is bound.</param>
    /// <param name="requestedProperties">The properties requested for this user configuration.</param>
    internal UserConfiguration(ExchangeService service, UserConfigurationProperties requestedProperties)
    {
        EwsUtilities.ValidateParam(service);

        if (service.RequestedServerVersion < ObjectVersion)
        {
            throw new ServiceVersionException(
                string.Format(Strings.ObjectTypeIncompatibleWithRequestVersion, GetType().Name, ObjectVersion)
            );
        }

        _service = service;
        _isNew = true;

        InitializeProperties(requestedProperties);
    }

    /// <summary>
    ///     Writes a byte array to Xml.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="byteArray">Byte array to write.</param>
    /// <param name="xmlElementName">Name of the Xml element.</param>
    private static void WriteByteArrayToXml(EwsServiceXmlWriter writer, byte[]? byteArray, string xmlElementName)
    {
        EwsUtilities.Assert(writer != null, "UserConfiguration.WriteByteArrayToXml", "writer is null");
        EwsUtilities.Assert(xmlElementName != null, "UserConfiguration.WriteByteArrayToXml", "xmlElementName is null");

        writer.WriteStartElement(XmlNamespace.Types, xmlElementName);

        if (byteArray != null && byteArray.Length > 0)
        {
            writer.WriteValue(Convert.ToBase64String(byteArray), xmlElementName);
        }

        writer.WriteEndElement();
    }

    /// <summary>
    ///     Writes to Xml.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="xmlNamespace">The XML namespace.</param>
    /// <param name="name">The user configuration name.</param>
    /// <param name="parentFolderId">The Id of the folder containing the user configuration.</param>
    internal static void WriteUserConfigurationNameToXml(
        EwsServiceXmlWriter writer,
        XmlNamespace xmlNamespace,
        string name,
        FolderId parentFolderId
    )
    {
        EwsUtilities.Assert(writer != null, "UserConfiguration.WriteUserConfigurationNameToXml", "writer is null");
        EwsUtilities.Assert(name != null, "UserConfiguration.WriteUserConfigurationNameToXml", "name is null");
        EwsUtilities.Assert(
            parentFolderId != null,
            "UserConfiguration.WriteUserConfigurationNameToXml",
            "parentFolderId is null"
        );

        writer.WriteStartElement(xmlNamespace, XmlElementNames.UserConfigurationName);

        writer.WriteAttributeValue(XmlAttributeNames.Name, name);

        parentFolderId.WriteToXml(writer);

        writer.WriteEndElement();
    }

    /// <summary>
    ///     Binds to an existing user configuration and loads the specified properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to which the user configuration is bound.</param>
    /// <param name="name">The name of the user configuration.</param>
    /// <param name="parentFolderId">The Id of the folder containing the user configuration.</param>
    /// <param name="properties">The properties to load.</param>
    /// <param name="token"></param>
    /// <returns>A user configuration instance.</returns>
    public static async Task<UserConfiguration> Bind(
        ExchangeService service,
        string name,
        FolderId parentFolderId,
        UserConfigurationProperties properties,
        CancellationToken token = default
    )
    {
        var result = await service.GetUserConfiguration(name, parentFolderId, properties, token).ConfigureAwait(false);

        result._isNew = false;

        return result;
    }

    /// <summary>
    ///     Binds to an existing user configuration and loads the specified properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to which the user configuration is bound.</param>
    /// <param name="name">The name of the user configuration.</param>
    /// <param name="parentFolderName">The name of the folder containing the user configuration.</param>
    /// <param name="properties">The properties to load.</param>
    /// <returns>A user configuration instance.</returns>
    public static Task<UserConfiguration> Bind(
        ExchangeService service,
        string name,
        WellKnownFolderName parentFolderName,
        UserConfigurationProperties properties
    )
    {
        return Bind(service, name, new FolderId(parentFolderName), properties);
    }

    /// <summary>
    ///     Saves the user configuration. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="name">The name of the user configuration.</param>
    /// <param name="parentFolderId">The Id of the folder in which to save the user configuration.</param>
    /// <param name="token"></param>
    public async System.Threading.Tasks.Task Save(
        string name,
        FolderId parentFolderId,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(name);
        EwsUtilities.ValidateParam(parentFolderId);

        parentFolderId.Validate(_service.RequestedServerVersion);

        if (!_isNew)
        {
            throw new InvalidOperationException(Strings.CannotSaveNotNewUserConfiguration);
        }

        ParentFolderId = parentFolderId;
        Name = name;

        await _service.CreateUserConfiguration(this, token).ConfigureAwait(false);

        _isNew = false;

        ResetIsDirty();
    }

    /// <summary>
    ///     Saves the user configuration. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="name">The name of the user configuration.</param>
    /// <param name="parentFolderName">The name of the folder in which to save the user configuration.</param>
    public System.Threading.Tasks.Task Save(string name, WellKnownFolderName parentFolderName)
    {
        return Save(name, new FolderId(parentFolderName));
    }

    /// <summary>
    ///     Updates the user configuration by applying local changes to the Exchange server.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    public async System.Threading.Tasks.Task Update(CancellationToken token = default)
    {
        if (_isNew)
        {
            throw new InvalidOperationException(Strings.CannotUpdateNewUserConfiguration);
        }

        if (IsPropertyUpdated(UserConfigurationProperties.BinaryData) ||
            IsPropertyUpdated(UserConfigurationProperties.Dictionary) ||
            IsPropertyUpdated(UserConfigurationProperties.XmlData))
        {
            await _service.UpdateUserConfiguration(this, token).ConfigureAwait(false);
        }

        ResetIsDirty();
    }

    /// <summary>
    ///     Deletes the user configuration. Calling this method results in a call to EWS.
    /// </summary>
    public async System.Threading.Tasks.Task Delete(CancellationToken token = default)
    {
        if (_isNew)
        {
            throw new InvalidOperationException(Strings.DeleteInvalidForUnsavedUserConfiguration);
        }

        await _service.DeleteUserConfiguration(Name, ParentFolderId, token).ConfigureAwait(false);
    }

    /// <summary>
    ///     Loads the specified properties on the user configuration. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="properties">The properties to load.</param>
    /// <param name="token"></param>
    public System.Threading.Tasks.Task Load(UserConfigurationProperties properties, CancellationToken token = default)
    {
        InitializeProperties(properties);

        return _service.LoadPropertiesForUserConfiguration(this, properties, token);
    }

    /// <summary>
    ///     Writes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="xmlNamespace">The XML namespace.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    internal void WriteToXml(EwsServiceXmlWriter writer, XmlNamespace xmlNamespace, string xmlElementName)
    {
        EwsUtilities.Assert(writer != null, "UserConfiguration.WriteToXml", "writer is null");
        EwsUtilities.Assert(xmlElementName != null, "UserConfiguration.WriteToXml", "xmlElementName is null");

        writer.WriteStartElement(xmlNamespace, xmlElementName);

        // Write the UserConfigurationName element
        WriteUserConfigurationNameToXml(writer, XmlNamespace.Types, Name, ParentFolderId);

        // Write the Dictionary element
        if (IsPropertyUpdated(UserConfigurationProperties.Dictionary))
        {
            Dictionary.WriteToXml(writer, XmlElementNames.Dictionary);
        }

        // Write the XmlData element
        if (IsPropertyUpdated(UserConfigurationProperties.XmlData))
        {
            WriteXmlDataToXml(writer);
        }

        // Write the BinaryData element
        if (IsPropertyUpdated(UserConfigurationProperties.BinaryData))
        {
            WriteBinaryDataToXml(writer);
        }

        writer.WriteEndElement();
    }

    /// <summary>
    ///     Determines whether the specified property was updated.
    /// </summary>
    /// <param name="property">property to evaluate.</param>
    /// <returns>Boolean indicating whether to send the property Xml.</returns>
    private bool IsPropertyUpdated(UserConfigurationProperties property)
    {
        var isPropertyDirty = false;
        var isPropertyEmpty = false;

        switch (property)
        {
            case UserConfigurationProperties.Dictionary:
            {
                isPropertyDirty = Dictionary.IsDirty;
                isPropertyEmpty = Dictionary.Count == 0;
                break;
            }
            case UserConfigurationProperties.XmlData:
            {
                isPropertyDirty = (property & _updatedProperties) == property;
                isPropertyEmpty = _xmlData == null || _xmlData.Length == 0;
                break;
            }
            case UserConfigurationProperties.BinaryData:
            {
                isPropertyDirty = (property & _updatedProperties) == property;
                isPropertyEmpty = _binaryData == null || _binaryData.Length == 0;
                break;
            }
            default:
            {
                EwsUtilities.Assert(
                    false,
                    "UserConfiguration.IsPropertyUpdated",
                    "property not supported: " + property
                );
                break;
            }
        }

        // Consider the property updated, if it's been modified, and either 
        //    . there's a value or 
        //    . there's no value but the operation is update.
        return isPropertyDirty && (!isPropertyEmpty || !_isNew);
    }

    /// <summary>
    ///     Writes the XmlData property to Xml.
    /// </summary>
    /// <param name="writer">The writer.</param>
    private void WriteXmlDataToXml(EwsServiceXmlWriter writer)
    {
        EwsUtilities.Assert(writer != null, "UserConfiguration.WriteXmlDataToXml", "writer is null");

        WriteByteArrayToXml(writer, _xmlData, XmlElementNames.XmlData);
    }

    /// <summary>
    ///     Writes the BinaryData property to Xml.
    /// </summary>
    /// <param name="writer">The writer.</param>
    private void WriteBinaryDataToXml(EwsServiceXmlWriter writer)
    {
        EwsUtilities.Assert(writer != null, "UserConfiguration.WriteBinaryDataToXml", "writer is null");

        WriteByteArrayToXml(writer, _binaryData, XmlElementNames.BinaryData);
    }

    /// <summary>
    ///     Loads from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal void LoadFromXml(EwsServiceXmlReader reader)
    {
        EwsUtilities.Assert(reader != null, "UserConfiguration.LoadFromXml", "reader is null");

        reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.UserConfiguration);
        reader.Read(); // Position at first property element

        do
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                switch (reader.LocalName)
                {
                    case XmlElementNames.UserConfigurationName:
                    {
                        var responseName = reader.ReadAttributeValue(XmlAttributeNames.Name);

                        EwsUtilities.Assert(
                            string.Compare(Name, responseName, StringComparison.Ordinal) == 0,
                            "UserConfiguration.LoadFromXml",
                            "UserConfigurationName does not match: Expected: " +
                            Name +
                            " Name in response: " +
                            responseName
                        );

                        reader.SkipCurrentElement();
                        break;
                    }
                    case XmlElementNames.ItemId:
                    {
                        ItemId = new ItemId();
                        ItemId.LoadFromXml(reader, XmlElementNames.ItemId);
                        break;
                    }
                    case XmlElementNames.Dictionary:
                    {
                        Dictionary.LoadFromXml(reader, XmlElementNames.Dictionary);
                        break;
                    }
                    case XmlElementNames.XmlData:
                    {
                        _xmlData = Convert.FromBase64String(reader.ReadElementValue());
                        break;
                    }
                    case XmlElementNames.BinaryData:
                    {
                        _binaryData = Convert.FromBase64String(reader.ReadElementValue());
                        break;
                    }
                    default:
                    {
                        EwsUtilities.Assert(
                            false,
                            "UserConfiguration.LoadFromXml",
                            "Xml element not supported: " + reader.LocalName
                        );
                        break;
                    }
                }
            }

            // If XmlData was loaded, read is skipped because GetXmlData positions the reader at the next property.
            reader.Read();
        } while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.UserConfiguration));
    }

    /// <summary>
    ///     Initializes properties.
    /// </summary>
    /// <param name="requestedProperties">The properties requested for this UserConfiguration.</param>
    /// <remarks>
    ///     InitializeProperties is called in 3 cases:
    ///     .  Create new object:  From the UserConfiguration constructor.
    ///     .  Bind to existing object:  Again from the constructor.  The constructor is called eventually by the
    ///     GetUserConfiguration request.
    ///     .  Refresh properties:  From the Load method.
    /// </remarks>
    private void InitializeProperties(UserConfigurationProperties requestedProperties)
    {
        ItemId = null;
        Dictionary = new UserConfigurationDictionary();
        _xmlData = null;
        _binaryData = null;
        _propertiesAvailableForAccess = requestedProperties;

        ResetIsDirty();
    }

    /// <summary>
    ///     Resets flags to indicate that properties haven't been modified.
    /// </summary>
    private void ResetIsDirty()
    {
        _updatedProperties = NoProperties;
        Dictionary.IsDirty = false;
    }

    /// <summary>
    ///     Determines whether the specified property may be accessed.
    /// </summary>
    /// <param name="property">Property to access.</param>
    private void ValidatePropertyAccess(UserConfigurationProperties property)
    {
        if ((property & _propertiesAvailableForAccess) != property)
        {
            throw new PropertyException(Strings.MustLoadOrAssignPropertyBeforeAccess, property.ToString());
        }
    }

    /// <summary>
    ///     Adds the passed property to updatedProperties.
    /// </summary>
    /// <param name="property">Property to update.</param>
    private void MarkPropertyForUpdate(UserConfigurationProperties property)
    {
        _updatedProperties |= property;
        _propertiesAvailableForAccess |= property;
    }
}
