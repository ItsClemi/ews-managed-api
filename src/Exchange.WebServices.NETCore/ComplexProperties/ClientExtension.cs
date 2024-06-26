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
///     Represents a ClientExtension object.
/// </summary>
[PublicAPI]
public sealed class ClientExtension : ComplexProperty
{
    /// <summary>
    ///     Gets or sets the extension type.
    /// </summary>
    public ExtensionType Type { get; set; }

    /// <summary>
    ///     Gets or sets the extension scope.
    /// </summary>
    public ExtensionInstallScope Scope { get; set; }

    /// <summary>
    ///     Gets or sets the extension manifest stream.
    /// </summary>
    public Stream? ManifestStream { get; set; }

    /// <summary>
    ///     Gets or sets the asset ID for Office Marketplace.
    /// </summary>
    public string MarketplaceAssetID { get; set; }

    /// <summary>
    ///     Gets or sets the content market for Office Marketplace.
    /// </summary>
    public string MarketplaceContentMarket { get; set; }

    /// <summary>
    ///     Gets or sets the app status
    /// </summary>
    public string AppStatus { get; set; }

    /// <summary>
    ///     Gets or sets the etoken
    /// </summary>
    public string Etoken { get; set; }

    /// <summary>
    ///     Gets or sets the Installed DateTime of the manifest
    /// </summary>
    public string InstalledDateTime { get; set; }

    /// <summary>
    ///     Gets or sets the value indicating whether extension is available.
    /// </summary>
    public bool IsAvailable { get; set; }

    /// <summary>
    ///     Gets or sets the value indicating whether extension is available.
    /// </summary>
    public bool IsMandatory { get; set; }

    /// <summary>
    ///     Gets or sets the value indicating whether extension is enabled by default.
    /// </summary>
    public bool IsEnabledByDefault { get; set; }

    /// <summary>
    ///     Gets or sets the extension ProvidedTo value.
    /// </summary>
    public ClientExtensionProvidedTo ProvidedTo { get; set; }

    /// <summary>
    ///     Gets or sets the user list this extension is provided to.
    /// </summary>
    public StringList? SpecificUsers { get; set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ClientExtension" /> class.
    /// </summary>
    internal ClientExtension()
    {
        Namespace = XmlNamespace.Types;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ClientExtension" /> class.
    /// </summary>
    /// <param name="type">Extension type</param>
    /// <param name="scope">Extension install scope</param>
    /// <param name="manifestStream">Manifest stream, can be null</param>
    /// <param name="marketplaceAssetId">The asset ID for Office Marketplace</param>
    /// <param name="marketplaceContentMarket">The content market for Office Marketplace</param>
    /// <param name="isAvailable">Whether extension is available</param>
    /// <param name="isMandatory">Whether extension is mandatory</param>
    /// <param name="isEnabledByDefault">Whether extension is enabled by default</param>
    /// <param name="providedTo">Who the extension is provided for (e.g. "entire org" or "specific users")</param>
    /// <param name="specificUsers">List of users extension is provided for, can be null</param>
    /// <param name="appStatus">App status</param>
    /// <param name="etoken">Etoken</param>
    public ClientExtension(
        ExtensionType type,
        ExtensionInstallScope scope,
        Stream manifestStream,
        string marketplaceAssetId,
        string marketplaceContentMarket,
        bool isAvailable,
        bool isMandatory,
        bool isEnabledByDefault,
        ClientExtensionProvidedTo providedTo,
        StringList specificUsers,
        string appStatus,
        string etoken
    )
        : this()
    {
        Type = type;
        Scope = scope;
        ManifestStream = manifestStream;
        MarketplaceAssetID = marketplaceAssetId;
        MarketplaceContentMarket = marketplaceContentMarket;
        IsAvailable = isAvailable;
        IsMandatory = isMandatory;
        IsEnabledByDefault = isEnabledByDefault;
        ProvidedTo = providedTo;
        SpecificUsers = specificUsers;
        AppStatus = appStatus;
        Etoken = etoken;
    }

    /// <summary>
    ///     Reads attributes from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
    {
        var value = reader.ReadAttributeValue(XmlAttributeNames.ClientExtensionType);
        if (!string.IsNullOrEmpty(value))
        {
            Type = reader.ReadAttributeValue<ExtensionType>(XmlAttributeNames.ClientExtensionType);
        }

        value = reader.ReadAttributeValue(XmlAttributeNames.ClientExtensionScope);
        if (!string.IsNullOrEmpty(value))
        {
            Scope = reader.ReadAttributeValue<ExtensionInstallScope>(XmlAttributeNames.ClientExtensionScope);
        }

        value = reader.ReadAttributeValue(XmlAttributeNames.ClientExtensionMarketplaceAssetID);
        if (!string.IsNullOrEmpty(value))
        {
            MarketplaceAssetID = reader.ReadAttributeValue<string>(XmlAttributeNames.ClientExtensionMarketplaceAssetID);
        }

        value = reader.ReadAttributeValue(XmlAttributeNames.ClientExtensionMarketplaceContentMarket);
        if (!string.IsNullOrEmpty(value))
        {
            MarketplaceContentMarket =
                reader.ReadAttributeValue<string>(XmlAttributeNames.ClientExtensionMarketplaceContentMarket);
        }

        value = reader.ReadAttributeValue(XmlAttributeNames.ClientExtensionAppStatus);
        if (!string.IsNullOrEmpty(value))
        {
            AppStatus = reader.ReadAttributeValue<string>(XmlAttributeNames.ClientExtensionAppStatus);
        }

        value = reader.ReadAttributeValue(XmlAttributeNames.ClientExtensionEtoken);
        if (!string.IsNullOrEmpty(value))
        {
            Etoken = reader.ReadAttributeValue<string>(XmlAttributeNames.ClientExtensionEtoken);
        }

        value = reader.ReadAttributeValue(XmlAttributeNames.ClientExtensionInstalledDateTime);
        if (!string.IsNullOrEmpty(value))
        {
            InstalledDateTime = reader.ReadAttributeValue<string>(XmlAttributeNames.ClientExtensionInstalledDateTime);
        }

        value = reader.ReadAttributeValue(XmlAttributeNames.ClientExtensionIsAvailable);
        if (!string.IsNullOrEmpty(value))
        {
            IsAvailable = reader.ReadAttributeValue<bool>(XmlAttributeNames.ClientExtensionIsAvailable);
        }

        value = reader.ReadAttributeValue(XmlAttributeNames.ClientExtensionIsMandatory);
        if (!string.IsNullOrEmpty(value))
        {
            IsMandatory = reader.ReadAttributeValue<bool>(XmlAttributeNames.ClientExtensionIsMandatory);
        }

        value = reader.ReadAttributeValue(XmlAttributeNames.ClientExtensionIsEnabledByDefault);
        if (!string.IsNullOrEmpty(value))
        {
            IsEnabledByDefault = reader.ReadAttributeValue<bool>(XmlAttributeNames.ClientExtensionIsEnabledByDefault);
        }

        value = reader.ReadAttributeValue(XmlAttributeNames.ClientExtensionProvidedTo);
        if (!string.IsNullOrEmpty(value))
        {
            ProvidedTo =
                reader.ReadAttributeValue<ClientExtensionProvidedTo>(XmlAttributeNames.ClientExtensionProvidedTo);
        }
    }

    /// <summary>
    ///     Writes attributes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionType, Type);
        writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionScope, Scope);
        writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionMarketplaceAssetID, MarketplaceAssetID);
        writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionMarketplaceContentMarket, MarketplaceContentMarket);
        writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionAppStatus, AppStatus);
        writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionEtoken, Etoken);
        writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionInstalledDateTime, InstalledDateTime);
        writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionIsAvailable, IsAvailable);
        writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionIsMandatory, IsMandatory);
        writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionIsEnabledByDefault, IsEnabledByDefault);
        writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionProvidedTo, ProvidedTo);
    }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.Manifest:
            {
                ManifestStream = new MemoryStream();
                reader.ReadBase64ElementValue(ManifestStream);
                ManifestStream.Position = 0;
                return true;
            }
            case XmlElementNames.ClientExtensionSpecificUsers:
            {
                SpecificUsers = new StringList();
                SpecificUsers.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.ClientExtensionSpecificUsers);
                return true;
            }
            default:
            {
                return base.TryReadElementFromXml(reader);
            }
        }
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        if (SpecificUsers != null)
        {
            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.ClientExtensionSpecificUsers);
            SpecificUsers.WriteElementsToXml(writer);
            writer.WriteEndElement();
        }

        if (ManifestStream != null)
        {
            if (ManifestStream.CanSeek)
            {
                ManifestStream.Position = 0;
            }

            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Manifest);
            writer.WriteBase64ElementValue(ManifestStream);
            writer.WriteEndElement();
        }
    }
}
