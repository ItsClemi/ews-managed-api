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

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a collection of ContactEntity objects.
/// </summary>
[PublicAPI]
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class ContactEntityCollection : ComplexPropertyCollection<ContactEntity>
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="ContactEntityCollection" /> class.
    /// </summary>
    internal ContactEntityCollection()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ContactEntityCollection" /> class.
    /// </summary>
    /// <param name="collection">The collection of objects to include.</param>
    internal ContactEntityCollection(IEnumerable<ContactEntity>? collection)
    {
        collection?.ForEach(InternalAdd);
    }

    /// <summary>
    ///     Creates the complex property.
    /// </summary>
    /// <param name="xmlElementName">Name of the XML element.</param>
    /// <returns>ContactEntity.</returns>
    internal override ContactEntity CreateComplexProperty(string xmlElementName)
    {
        return new ContactEntity();
    }

    /// <summary>
    ///     Gets the name of the collection item XML element.
    /// </summary>
    /// <param name="complexProperty">The complex property.</param>
    /// <returns>XML element name.</returns>
    internal override string GetCollectionItemXmlElementName(ContactEntity complexProperty)
    {
        return XmlElementNames.NlgContact;
    }
}
