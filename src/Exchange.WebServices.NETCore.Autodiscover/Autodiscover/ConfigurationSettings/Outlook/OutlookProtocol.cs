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
using System.Xml;

using Microsoft.Exchange.WebServices.Data;

namespace Microsoft.Exchange.WebServices.Autodiscover;

/// <summary>
///     Represents a supported Outlook protocol in an Outlook configurations settings account.
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
internal sealed class OutlookProtocol
{
    private const string EXCH = "EXCH";
    private const string EXPR = "EXPR";
    private const string WEB = "WEB";


    #region Private static fields

    /// <summary>
    ///     Converters to translate common Outlook protocol settings.
    ///     Each entry maps to a lambda expression used to get the matching property from the OutlookProtocol instance.
    /// </summary>
    private static readonly IReadOnlyDictionary<UserSettingName, Func<OutlookProtocol, object>> CommonProtocolSettings =
        new Dictionary<UserSettingName, Func<OutlookProtocol, object>>
        {
            // @formatter:off
            { UserSettingName.EcpDeliveryReportUrlFragment, p => p._ecpUrlMt },
            { UserSettingName.EcpEmailSubscriptionsUrlFragment, p => p._ecpUrlAggr },
            { UserSettingName.EcpPublishingUrlFragment, p => p._ecpUrlPublish },
            { UserSettingName.EcpPhotoUrlFragment, p => p._ecpUrlPhoto },
            { UserSettingName.EcpRetentionPolicyTagsUrlFragment, p => p._ecpUrlRet },
            { UserSettingName.EcpTextMessagingUrlFragment, p => p._ecpUrlSms },
            { UserSettingName.EcpVoicemailUrlFragment, p => p._ecpUrlUm },
            { UserSettingName.EcpConnectUrlFragment, p => p._ecpUrlConnect },
            { UserSettingName.EcpTeamMailboxUrlFragment, p => p._ecpUrlTm },
            { UserSettingName.EcpTeamMailboxCreatingUrlFragment, p => p._ecpUrlTmCreating },
            { UserSettingName.EcpTeamMailboxEditingUrlFragment, p => p._ecpUrlTmEditing },
            { UserSettingName.EcpExtensionInstallationUrlFragment, p => p._ecpUrlExtInstall },
            { UserSettingName.SiteMailboxCreationURL, p => p._siteMailboxCreationUrl },
            // @formatter:on
        };

    /// <summary>
    ///     Converters to translate internal (EXCH) Outlook protocol settings.
    ///     Each entry maps to a lambda expression used to get the matching property from the OutlookProtocol instance.
    /// </summary>
    private static readonly IReadOnlyDictionary<UserSettingName, Func<OutlookProtocol, object?>>
        InternalProtocolSettings = new Dictionary<UserSettingName, Func<OutlookProtocol, object?>>
        {
            // @formatter:off
            { UserSettingName.ActiveDirectoryServer, p => p._activeDirectoryServer },
            { UserSettingName.CrossOrganizationSharingEnabled, p => p._sharingEnabled.ToString() },
            { UserSettingName.InternalEcpUrl, p => p._ecpUrl },
            { UserSettingName.InternalEcpDeliveryReportUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlMt) },
            { UserSettingName.InternalEcpEmailSubscriptionsUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlAggr) },
            { UserSettingName.InternalEcpPublishingUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlPublish) },
            { UserSettingName.InternalEcpPhotoUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlPhoto) },
            { UserSettingName.InternalEcpRetentionPolicyTagsUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlRet) },
            { UserSettingName.InternalEcpTextMessagingUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlSms) },
            { UserSettingName.InternalEcpVoicemailUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlUm) },
            { UserSettingName.InternalEcpConnectUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlConnect) },
            { UserSettingName.InternalEcpTeamMailboxUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlTm) },
            { UserSettingName.InternalEcpTeamMailboxCreatingUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlTmCreating) },
            { UserSettingName.InternalEcpTeamMailboxEditingUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlTmEditing) },
            { UserSettingName.InternalEcpTeamMailboxHidingUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlTmHiding) },
            { UserSettingName.InternalEcpExtensionInstallationUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlExtInstall) },
            { UserSettingName.InternalEwsUrl, p => p._exchangeWebServicesUrl ?? p._availabilityServiceUrl },
            { UserSettingName.InternalEmwsUrl, p => p._exchangeManagementWebServicesUrl },
            { UserSettingName.InternalMailboxServerDN, p => p._serverDn },
            { UserSettingName.InternalRpcClientServer, p => p._server },
            { UserSettingName.InternalOABUrl, p => p._offlineAddressBookUrl },
            { UserSettingName.InternalUMUrl, p => p._unifiedMessagingUrl },
            { UserSettingName.MailboxDN, p => p._mailboxDn },
            { UserSettingName.PublicFolderServer, p => p._publicFolderServer },
            { UserSettingName.InternalServerExclusiveConnect, p => p._serverExclusiveConnect },
            // @formatter:on
        };

    /// <summary>
    ///     Converters to translate external (EXPR) Outlook protocol settings.
    ///     Each entry maps to a lambda expression used to get the matching property from the OutlookProtocol instance.
    /// </summary>
    private static readonly IReadOnlyDictionary<UserSettingName, Func<OutlookProtocol, object?>>
        ExternalProtocolSettings = new Dictionary<UserSettingName, Func<OutlookProtocol, object?>>
        {
            // @formatter:off
            { UserSettingName.ExternalEcpDeliveryReportUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlRet) },
            { UserSettingName.ExternalEcpEmailSubscriptionsUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlAggr) },
            { UserSettingName.ExternalEcpPublishingUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlPublish) },
            { UserSettingName.ExternalEcpPhotoUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlPhoto) },
            { UserSettingName.ExternalEcpRetentionPolicyTagsUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlRet) },
            { UserSettingName.ExternalEcpTextMessagingUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlSms) },
            { UserSettingName.ExternalEcpUrl, p => p._ecpUrl },
            { UserSettingName.ExternalEcpVoicemailUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlUm) },
            { UserSettingName.ExternalEcpConnectUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlConnect) },
            { UserSettingName.ExternalEcpTeamMailboxUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlTm) },
            { UserSettingName.ExternalEcpTeamMailboxCreatingUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlTmCreating) },
            { UserSettingName.ExternalEcpTeamMailboxEditingUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlTmEditing) },
            { UserSettingName.ExternalEcpTeamMailboxHidingUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlTmHiding) },
            { UserSettingName.ExternalEcpExtensionInstallationUrl, p => p.ConvertEcpFragmentToUrl(p._ecpUrlExtInstall) },
            { UserSettingName.ExternalEwsUrl, p => p._exchangeWebServicesUrl ?? p._availabilityServiceUrl },
            { UserSettingName.ExternalEmwsUrl, p => p._exchangeManagementWebServicesUrl },
            { UserSettingName.ExternalMailboxServer, p => p._server },
            { UserSettingName.ExternalMailboxServerAuthenticationMethods, p => p._authPackage },
            { UserSettingName.ExternalMailboxServerRequiresSSL, p => p._sslEnabled.ToString() },
            { UserSettingName.ExternalOABUrl, p => p._offlineAddressBookUrl },
            { UserSettingName.ExternalUMUrl, p => p._unifiedMessagingUrl },
            { UserSettingName.ExchangeRpcUrl, p => p._exchangeRpcUrl },
            { UserSettingName.EwsPartnerUrl, p => p._exchangeWebServicesPartnerUrl },
            { UserSettingName.ExternalServerExclusiveConnect, p => p._serverExclusiveConnect.ToString() },
            { UserSettingName.CertPrincipalName, p => p._certPrincipalName },
            { UserSettingName.GroupingInformation, p => p._groupingInformation },
            // @formatter:on
        };

    /// <summary>
    ///     Merged converter dictionary for translating internal (EXCH) Outlook protocol settings.
    ///     Each entry maps to a lambda expression used to get the matching property from the OutlookProtocol instance.
    /// </summary>
    private static readonly Lazy<Dictionary<UserSettingName, Func<OutlookProtocol, object?>>>
        InternalProtocolConverterDictionary = new(
            () =>
            {
                var results = new Dictionary<UserSettingName, Func<OutlookProtocol, object?>>();

                foreach (var (key, value) in CommonProtocolSettings.ToList())
                {
                    results.Add(key, value);
                }

                foreach (var kv in InternalProtocolSettings.ToList())
                {
                    results.Add(kv.Key, kv.Value);
                }

                return results;
            }
        );

    /// <summary>
    ///     Merged converter dictionary for translating external (EXPR) Outlook protocol settings.
    ///     Each entry maps to a lambda expression used to get the matching property from the OutlookProtocol instance.
    /// </summary>
    private static readonly Lazy<Dictionary<UserSettingName, Func<OutlookProtocol, object?>>>
        ExternalProtocolConverterDictionary = new(
            () =>
            {
                var results = new Dictionary<UserSettingName, Func<OutlookProtocol, object?>>();

                foreach (var (key, value) in CommonProtocolSettings.ToList())
                {
                    results.Add(key, value);
                }

                foreach (var (key, value) in ExternalProtocolSettings.ToList())
                {
                    results.Add(key, value);
                }

                return results;
            }
        );

    /// <summary>
    ///     Converters to translate Web (WEB) Outlook protocol settings.
    ///     Each entry maps to a lambda expression used to get the matching property from the OutlookProtocol instance.
    /// </summary>
    private static readonly IReadOnlyDictionary<UserSettingName, Func<OutlookProtocol, object>>
        WebProtocolConverterDictionary = new Dictionary<UserSettingName, Func<OutlookProtocol, object>>
        {
            // @formatter:off
            { UserSettingName.InternalWebClientUrls, p => p._internalOutlookWebAccessUrls },
            { UserSettingName.ExternalWebClientUrls, p => p._externalOutlookWebAccessUrls },
            // @formatter:on
        };

    /// <summary>
    ///     The collection of available user settings for all OutlookProtocol types.
    /// </summary>
    private static readonly Lazy<List<UserSettingName>> availableUserSettings = new(
        () =>
        {
            var results = new List<UserSettingName>();
            results.AddRange(CommonProtocolSettings.Keys);
            results.AddRange(InternalProtocolSettings.Keys);
            results.AddRange(ExternalProtocolSettings.Keys);
            results.AddRange(WebProtocolConverterDictionary.Keys);
            return results;
        }
    );

    /// <summary>
    ///     Map Outlook protocol name to type.
    /// </summary>
    private static readonly IReadOnlyDictionary<string, OutlookProtocolType> ProtocolNameToTypeMap =
        new Dictionary<string, OutlookProtocolType>
        {
            // @formatter:off
            { EXCH, OutlookProtocolType.Rpc },
            { EXPR, OutlookProtocolType.RpcOverHttp },
            { WEB, OutlookProtocolType.Web },
            // @formatter:on
        };

    #endregion


    #region Private fields

    private string _activeDirectoryServer;
    private string _authPackage;
    private string _availabilityServiceUrl;
    private string _ecpUrl;
    private string _ecpUrlAggr;
    private string _ecpUrlMt;
    private string _ecpUrlPublish;
    private string _ecpUrlPhoto;
    private string _ecpUrlConnect;
    private string _ecpUrlRet;
    private string _ecpUrlSms;
    private string _ecpUrlUm;
    private string _ecpUrlTm;
    private string _ecpUrlTmCreating;
    private string _ecpUrlTmEditing;
    private string _ecpUrlTmHiding;
    private string _siteMailboxCreationUrl;
    private string _ecpUrlExtInstall;
    private string? _exchangeWebServicesUrl;
    private string _exchangeManagementWebServicesUrl;
    private string _mailboxDn;
    private string _offlineAddressBookUrl;
    private string _exchangeRpcUrl;
    private string _exchangeWebServicesPartnerUrl;
    private string _publicFolderServer;
    private string _server;
    private string _serverDn;
    private string _unifiedMessagingUrl;
    private bool _sharingEnabled;
    private bool _sslEnabled;
    private bool _serverExclusiveConnect;
    private string _certPrincipalName;
    private string _groupingInformation;
    private readonly WebClientUrlCollection _externalOutlookWebAccessUrls;
    private readonly WebClientUrlCollection _internalOutlookWebAccessUrls;

    #endregion


    /// <summary>
    ///     Initializes a new instance of the <see cref="OutlookProtocol" /> class.
    /// </summary>
    internal OutlookProtocol()
    {
        _internalOutlookWebAccessUrls = new WebClientUrlCollection();
        _externalOutlookWebAccessUrls = new WebClientUrlCollection();
    }

    /// <summary>
    ///     Loads from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    internal void LoadFromXml(EwsXmlReader reader)
    {
        do
        {
            reader.Read();

            if (reader.NodeType == XmlNodeType.Element)
            {
                switch (reader.LocalName)
                {
                    case XmlElementNames.Type:
                    {
                        ProtocolType = ProtocolNameToType(reader.ReadElementValue());
                        break;
                    }
                    case XmlElementNames.AuthPackage:
                    {
                        _authPackage = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.Server:
                    {
                        _server = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.ServerDN:
                    {
                        _serverDn = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.ServerVersion:
                    {
                        // just read it out
                        reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.AD:
                    {
                        _activeDirectoryServer = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.MdbDN:
                    {
                        _mailboxDn = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.EWSUrl:
                    {
                        _exchangeWebServicesUrl = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.EmwsUrl:
                    {
                        _exchangeManagementWebServicesUrl = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.ASUrl:
                    {
                        _availabilityServiceUrl = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.OOFUrl:
                    {
                        // just read it out
                        reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.UMUrl:
                    {
                        _unifiedMessagingUrl = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.OABUrl:
                    {
                        _offlineAddressBookUrl = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.PublicFolderServer:
                    {
                        _publicFolderServer = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.Internal:
                    {
                        LoadWebClientUrlsFromXml(reader, _internalOutlookWebAccessUrls, reader.LocalName);
                        break;
                    }
                    case XmlElementNames.External:
                    {
                        LoadWebClientUrlsFromXml(reader, _externalOutlookWebAccessUrls, reader.LocalName);
                        break;
                    }
                    case XmlElementNames.Ssl:
                    {
                        var sslStr = reader.ReadElementValue();
                        _sslEnabled = sslStr.Equals("On", StringComparison.OrdinalIgnoreCase);
                        break;
                    }
                    case XmlElementNames.SharingUrl:
                    {
                        _sharingEnabled = reader.ReadElementValue().Length > 0;
                        break;
                    }
                    case XmlElementNames.EcpUrl:
                    {
                        _ecpUrl = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.EcpUrl_um:
                    {
                        _ecpUrlUm = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.EcpUrl_aggr:
                    {
                        _ecpUrlAggr = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.EcpUrl_sms:
                    {
                        _ecpUrlSms = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.EcpUrl_mt:
                    {
                        _ecpUrlMt = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.EcpUrl_ret:
                    {
                        _ecpUrlRet = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.EcpUrl_publish:
                    {
                        _ecpUrlPublish = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.EcpUrl_photo:
                    {
                        _ecpUrlPhoto = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.ExchangeRpcUrl:
                    {
                        _exchangeRpcUrl = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.EwsPartnerUrl:
                    {
                        _exchangeWebServicesPartnerUrl = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.EcpUrl_connect:
                    {
                        _ecpUrlConnect = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.EcpUrl_tm:
                    {
                        _ecpUrlTm = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.EcpUrl_tmCreating:
                    {
                        _ecpUrlTmCreating = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.EcpUrl_tmEditing:
                    {
                        _ecpUrlTmEditing = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.EcpUrl_tmHiding:
                    {
                        _ecpUrlTmHiding = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.SiteMailboxCreationURL:
                    {
                        _siteMailboxCreationUrl = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.EcpUrl_extinstall:
                    {
                        _ecpUrlExtInstall = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.ServerExclusiveConnect:
                    {
                        var serverExclusiveConnectStr = reader.ReadElementValue();
                        _serverExclusiveConnect = serverExclusiveConnectStr.Equals(
                            "On",
                            StringComparison.OrdinalIgnoreCase
                        );
                        break;
                    }
                    case XmlElementNames.CertPrincipalName:
                    {
                        _certPrincipalName = reader.ReadElementValue();
                        break;
                    }
                    case XmlElementNames.GroupingInformation:
                    {
                        _groupingInformation = reader.ReadElementValue();
                        break;
                    }
                    default:
                    {
                        reader.SkipCurrentElement();
                        break;
                    }
                }
            }
        } while (!reader.IsEndElement(XmlNamespace.NotSpecified, XmlElementNames.Protocol));
    }

    /// <summary>
    ///     Convert protocol name to protocol type.
    /// </summary>
    /// <param name="protocolName">Name of the protocol.</param>
    /// <returns>OutlookProtocolType</returns>
    private static OutlookProtocolType ProtocolNameToType(string protocolName)
    {
        if (!ProtocolNameToTypeMap.TryGetValue(protocolName, out var protocolType))
        {
            protocolType = OutlookProtocolType.Unknown;
        }

        return protocolType;
    }

    /// <summary>
    ///     Loads web client urls from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="webClientUrls">The web client urls.</param>
    /// <param name="elementName">Name of the element.</param>
    private static void LoadWebClientUrlsFromXml(
        EwsXmlReader reader,
        WebClientUrlCollection webClientUrls,
        string elementName
    )
    {
        do
        {
            reader.Read();

            if (reader.NodeType == XmlNodeType.Element)
            {
                switch (reader.LocalName)
                {
                    case XmlElementNames.OWAUrl:
                    {
                        var authMethod = reader.ReadAttributeValue(XmlAttributeNames.AuthenticationMethod);
                        var owaUrl = reader.ReadElementValue();
                        var webClientUrl = new WebClientUrl(authMethod, owaUrl);
                        webClientUrls.Urls.Add(webClientUrl);
                        break;
                    }
                    default:
                    {
                        reader.SkipCurrentElement();
                        break;
                    }
                }
            }
        } while (!reader.IsEndElement(XmlNamespace.NotSpecified, elementName));
    }

    /// <summary>
    ///     Convert ECP fragment to full ECP URL.
    /// </summary>
    /// <param name="fragment">The fragment.</param>
    /// <returns>Full URL string (or null if either portion is empty.</returns>
    private string? ConvertEcpFragmentToUrl(string fragment)
    {
        return string.IsNullOrEmpty(_ecpUrl) || string.IsNullOrEmpty(fragment) ? null : _ecpUrl + fragment;
    }

    /// <summary>
    ///     Convert OutlookProtocol to GetUserSettings response.
    /// </summary>
    /// <param name="requestedSettings">The requested settings.</param>
    /// <param name="response">The response.</param>
    internal void ConvertToUserSettings(List<UserSettingName> requestedSettings, GetUserSettingsResponse response)
    {
        if (ConverterDictionary != null)
        {
            // In English: collect converters that are contained in the requested settings.
            var converterQuery = from converter in ConverterDictionary
                where requestedSettings.Contains(converter.Key)
                select converter;

            foreach (var kv in converterQuery)
            {
                var value = kv.Value(this);
                if (value != null)
                {
                    response.Settings[kv.Key] = value;
                }
            }
        }
    }

    /// <summary>
    ///     Gets the type of the protocol.
    /// </summary>
    /// <value>The type of the protocol.</value>
    internal OutlookProtocolType ProtocolType { get; set; }

    /// <summary>
    ///     Gets the converter dictionary for protocol type.
    /// </summary>
    /// <value>The converter dictionary.</value>
    private IReadOnlyDictionary<UserSettingName, Func<OutlookProtocol, object>>? ConverterDictionary
    {
        get =>
            ProtocolType switch
            {
                OutlookProtocolType.Rpc => InternalProtocolConverterDictionary.Value,
                OutlookProtocolType.RpcOverHttp => ExternalProtocolConverterDictionary.Value,
                OutlookProtocolType.Web => WebProtocolConverterDictionary,
                _ => null,
            };
    }

    /// <summary>
    ///     Gets the available user settings.
    /// </summary>
    internal static List<UserSettingName> AvailableUserSettings => availableUserSettings.Value;
}
