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

using Microsoft.Exchange.WebServices.Data;

namespace Microsoft.Exchange.WebServices.Autodiscover;

/// <summary>
///     Represents a collection of responses to GetDomainSettings
/// </summary>
[PublicAPI]
public sealed class GetDomainSettingsResponseCollection : AutodiscoverResponseCollection<GetDomainSettingsResponse>
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="AutodiscoverResponseCollection&lt;TResponse&gt;" /> class.
    /// </summary>
    internal GetDomainSettingsResponseCollection()
    {
    }

    /// <summary>
    ///     Create a response instance.
    /// </summary>
    /// <returns>GetDomainSettingsResponse.</returns>
    internal override GetDomainSettingsResponse CreateResponseInstance()
    {
        return new GetDomainSettingsResponse();
    }

    /// <summary>
    ///     Gets the name of the response collection XML element.
    /// </summary>
    /// <returns>Response collection XMl element name.</returns>
    internal override string GetResponseCollectionXmlElementName()
    {
        return XmlElementNames.DomainResponses;
    }

    /// <summary>
    ///     Gets the name of the response instance XML element.
    /// </summary>
    /// <returns>Response instance XMl element name.</returns>
    internal override string GetResponseInstanceXmlElementName()
    {
        return XmlElementNames.DomainResponse;
    }
}
