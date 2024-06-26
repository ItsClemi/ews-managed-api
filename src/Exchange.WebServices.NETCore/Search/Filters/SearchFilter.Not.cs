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

/// <content>
///     Contains nested type SearchFilter.Not.
/// </content>
public abstract partial class SearchFilter
{
    /// <summary>
    ///     Represents a search filter that negates another. Applications can use NotFilter to define
    ///     conditions such as "NOT(other filter)".
    /// </summary>
    [PublicAPI]
    public sealed class Not : SearchFilter
    {
        private SearchFilter? _searchFilter;

        /// <summary>
        ///     Gets or sets the search filter to negate. Available search filter classes include
        ///     SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and SearchFilter.SearchFilterCollection.
        /// </summary>
        public SearchFilter? SearchFilter
        {
            get => _searchFilter;

            set
            {
                if (_searchFilter != null)
                {
                    _searchFilter.OnChange -= SearchFilterChanged;
                }

                SetFieldValue(ref _searchFilter, value);

                if (_searchFilter != null)
                {
                    _searchFilter.OnChange += SearchFilterChanged;
                }
            }
        }

        /// <summary>
        ///     Initializes a new instance of the <see cref="SearchFilter.Not" /> class.
        /// </summary>
        public Not()
        {
        }

        /// <summary>
        ///     Initializes a new instance of the <see cref="SearchFilter.Not" /> class.
        /// </summary>
        /// <param name="searchFilter">
        ///     The search filter to negate. Available search filter classes include SearchFilter.IsEqualTo,
        ///     SearchFilter.ContainsSubstring and SearchFilter.SearchFilterCollection.
        /// </param>
        public Not(SearchFilter searchFilter)
        {
            _searchFilter = searchFilter;
        }

        /// <summary>
        ///     A search filter has changed.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        private void SearchFilterChanged(ComplexProperty complexProperty)
        {
            Changed();
        }

        /// <summary>
        ///     Validate instance.
        /// </summary>
        internal override void InternalValidate()
        {
            if (_searchFilter == null)
            {
                throw new ServiceValidationException(Strings.SearchFilterMustBeSet);
            }
        }

        /// <summary>
        ///     Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.Not;
        }

        /// <summary>
        ///     Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            _searchFilter = LoadFromXml(reader);
            return true;
        }

        /// <summary>
        ///     Writes the elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            SearchFilter.WriteToXml(writer);
        }
    }
}
