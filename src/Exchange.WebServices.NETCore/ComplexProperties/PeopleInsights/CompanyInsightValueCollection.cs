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
///     Represents the collection of job insights.
/// </summary>
[PublicAPI]
public sealed class CompanyInsightValueCollection : ComplexPropertyCollection<CompanyInsightValue>
{
    /// <summary>
    ///     XML element name
    /// </summary>
    private readonly string _collectionItemXmlElementName;

    /// <summary>
    ///     Creates a new instance of the <see cref="CompanyInsightValueCollection" /> class.
    /// </summary>
    internal CompanyInsightValueCollection()
        : this(XmlElementNames.CompanyInsightValue)
    {
    }

    /// <summary>
    ///     Creates a new instance of the <see cref="CompanyInsightValueCollection" /> class.
    /// </summary>
    /// <param name="collectionItemXmlElementName">Name of the collection item XML element.</param>
    internal CompanyInsightValueCollection(string collectionItemXmlElementName)
    {
        _collectionItemXmlElementName = collectionItemXmlElementName;
    }

    /// <summary>
    ///     Creates a CompanyInsightValue object from an XML element name.
    /// </summary>
    /// <param name="xmlElementName">The XML element name from which to create the CompanyInsightValue.</param>
    /// <returns>A CompanyInsightValue object.</returns>
    internal override CompanyInsightValue? CreateComplexProperty(string xmlElementName)
    {
        if (xmlElementName == _collectionItemXmlElementName)
        {
            return new CompanyInsightValue();
        }

        return null;
    }

    /// <summary>
    ///     Retrieves the XML element name corresponding to the provided PersonInsight object.
    /// </summary>
    /// <param name="insight">The CompanyInsightValue object from which to determine the XML element name.</param>
    /// <returns>The XML element name corresponding to the provided CompanyInsightValue object.</returns>
    internal override string GetCollectionItemXmlElementName(CompanyInsightValue insight)
    {
        return _collectionItemXmlElementName;
    }

    /// <summary>
    ///     Determine whether we should write collection to XML or not.
    /// </summary>
    /// <returns>Always true, even if the collection is empty.</returns>
    internal override bool ShouldWriteToRequest()
    {
        return true;
    }
}
