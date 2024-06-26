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

using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Xml;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     X509CertificateCredentials wraps an instance of X509Certificate2, it can be used for WS-Security/X509
///     certificate-based authentication.
/// </summary>
[PublicAPI]
public sealed class X509CertificateCredentials : WSSecurityBasedCredentials
{
    private const string BinarySecurityTokenFormat = "<wsse:BinarySecurityToken " +
                                                     "EncodingType=\"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-soap-message-security-1.0#Base64Binary\" " +
                                                     "ValueType=\"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-x509-token-profile-1.0#X509v3\" " +
                                                     "wsu:Id=\"{0}\">" +
                                                     "{1}" +
                                                     "</wsse:BinarySecurityToken>";

    private const string KeyInfoClauseFormat =
        "<wsse:SecurityTokenReference xmlns:wsse=\"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd\" >" +
        "<wsse:Reference URI=\"#{0}\" ValueType=\"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-x509-token-profile-1.0#X509v3\" />" +
        "</wsse:SecurityTokenReference>";

    private const string WsSecurityX509CertPathSuffix = "/wssecurity/x509cert";

    private readonly X509Certificate2 _certificate;

    private readonly KeyInfoClause _keyInfoClause;

    /// <summary>
    ///     Gets the flag indicating whether any sign action need taken.
    /// </summary>
    internal override bool NeedSignature => true;

    /// <summary>
    ///     Initializes a new instance of the <see cref="X509CertificateCredentials" /> class.
    /// </summary>
    /// <remarks>The X509Certificate2 argument should have private key in order to sign the message.</remarks>
    /// <param name="certificate">The X509Certificate2 object.</param>
    public X509CertificateCredentials(X509Certificate2 certificate)
        : base(null, true)
    {
        EwsUtilities.ValidateParam(certificate);

        if (!certificate.HasPrivateKey)
        {
            throw new ServiceValidationException(Strings.CertificateHasNoPrivateKey);
        }

        _certificate = certificate;

        var certId = WsSecurityUtilityIdSignedXml.GetUniqueId();

        SecurityToken = string.Format(
            BinarySecurityTokenFormat,
            certId,
            Convert.ToBase64String(_certificate.GetRawCertData())
        );

        var doc = new SafeXmlDocument
        {
            PreserveWhitespace = true,
        };
        doc.LoadXml(string.Format(KeyInfoClauseFormat, certId));
        _keyInfoClause = new KeyInfoNode(doc.DocumentElement);
    }

    /// <summary>
    ///     This method is called to apply credentials to a service request before the request is made.
    /// </summary>
    /// <param name="request">The request.</param>
    internal override System.Threading.Tasks.Task PrepareWebRequest(EwsHttpWebRequest request)
    {
        EwsUrl = request.RequestUri;

        return System.Threading.Tasks.Task.CompletedTask;
    }

    /// <summary>
    ///     Adjusts the URL based on the credentials.
    /// </summary>
    /// <param name="url">The URL.</param>
    /// <returns>Adjust URL.</returns>
    internal override Uri AdjustUrl(Uri url)
    {
        return new Uri(GetUriWithoutSuffix(url) + WsSecurityX509CertPathSuffix);
    }

    /// <summary>
    ///     Add the signature element to the memory stream.
    /// </summary>
    /// <param name="memoryStream">The memory stream.</param>
    internal override void Sign(MemoryStream memoryStream)
    {
        memoryStream.Position = 0;

        var document = new SafeXmlDocument
        {
            PreserveWhitespace = true,
        };
        document.Load(memoryStream);

        var signedXml = new WsSecurityUtilityIdSignedXml(document)
        {
            SignedInfo =
            {
                CanonicalizationMethod = SignedXml.XmlDsigExcC14NTransformUrl,
            },
            SigningKey = _certificate.PrivateKey,
        };

        signedXml.AddReference("/soap:Envelope/soap:Header/wsa:To");
        signedXml.AddReference("/soap:Envelope/soap:Header/wsse:Security/wsu:Timestamp");

        signedXml.KeyInfo.AddClause(_keyInfoClause);
        signedXml.ComputeSignature();
        var signature = signedXml.GetXml();

        var wsSecurityNode = document.SelectSingleNode("/soap:Envelope/soap:Header/wsse:Security", NamespaceManager);

        wsSecurityNode.AppendChild(signature);

        memoryStream.Position = 0;
        document.Save(memoryStream);
    }

    /// <summary>
    ///     Gets the credentials string presentation.
    /// </summary>
    /// <returns>The string.</returns>
    public override string ToString()
    {
        return $"X509:<I>={_certificate.Issuer},<S>={_certificate.Subject}";
    }
}
