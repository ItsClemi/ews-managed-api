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

using System.Collections;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

using PropertyDefinitionSortDirectionPair = KeyValuePair<PropertyDefinitionBase, SortDirection>;

/// <summary>
///     Represents an ordered collection of property definitions qualified with a sort direction.
/// </summary>
[PublicAPI]
public sealed class OrderByCollection : IEnumerable<PropertyDefinitionSortDirectionPair>
{
    private readonly List<PropertyDefinitionSortDirectionPair> _propDefSortOrderPairList;

    /// <summary>
    ///     Gets the number of elements contained in the collection.
    /// </summary>
    public int Count => _propDefSortOrderPairList.Count;

    /// <summary>
    ///     Gets the element at the specified index from the collection.
    /// </summary>
    /// <param name="index">Index.</param>
    public PropertyDefinitionSortDirectionPair this[int index] => _propDefSortOrderPairList[index];

    /// <summary>
    ///     Initializes a new instance of the <see cref="OrderByCollection" /> class.
    /// </summary>
    internal OrderByCollection()
    {
        _propDefSortOrderPairList = new List<PropertyDefinitionSortDirectionPair>();
    }


    #region IEnumerable<KeyValuePair<PropertyDefinitionBase,SortDirection>> Members

    /// <summary>
    ///     Returns an enumerator that iterates through the collection.
    /// </summary>
    /// <returns>
    ///     A <see cref="T:System.Collections.Generic.IEnumerator`1" /> that can be used to iterate through the collection.
    /// </returns>
    public IEnumerator<KeyValuePair<PropertyDefinitionBase, SortDirection>> GetEnumerator()
    {
        return _propDefSortOrderPairList.GetEnumerator();
    }

    #endregion


    #region IEnumerable Members

    /// <summary>
    ///     Returns an enumerator that iterates through a collection.
    /// </summary>
    /// <returns>
    ///     An <see cref="T:System.Collections.IEnumerator" /> object that can be used to iterate through the collection.
    /// </returns>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return _propDefSortOrderPairList.GetEnumerator();
    }

    #endregion


    /// <summary>
    ///     Adds the specified property definition / sort direction pair to the collection.
    /// </summary>
    /// <param name="propertyDefinition">The property definition.</param>
    /// <param name="sortDirection">The sort direction.</param>
    public void Add(PropertyDefinitionBase propertyDefinition, SortDirection sortDirection)
    {
        if (Contains(propertyDefinition))
        {
            throw new ServiceLocalException(
                string.Format(Strings.PropertyAlreadyExistsInOrderByCollection, propertyDefinition.GetPrintableName())
            );
        }

        _propDefSortOrderPairList.Add(new PropertyDefinitionSortDirectionPair(propertyDefinition, sortDirection));
    }

    /// <summary>
    ///     Removes all elements from the collection.
    /// </summary>
    public void Clear()
    {
        _propDefSortOrderPairList.Clear();
    }

    /// <summary>
    ///     Determines whether the collection contains the specified property definition.
    /// </summary>
    /// <param name="propertyDefinition">The property definition.</param>
    /// <returns>True if the collection contains the specified property definition; otherwise, false.</returns>
    internal bool Contains(PropertyDefinitionBase propertyDefinition)
    {
        return _propDefSortOrderPairList.Exists(pair => pair.Key.Equals(propertyDefinition));
    }

    /// <summary>
    ///     Removes the specified property definition from the collection.
    /// </summary>
    /// <param name="propertyDefinition">The property definition.</param>
    /// <returns>True if the property definition is successfully removed; otherwise, false</returns>
    public bool Remove(PropertyDefinitionBase propertyDefinition)
    {
        var count = _propDefSortOrderPairList.RemoveAll(pair => pair.Key.Equals(propertyDefinition));
        return count > 0;
    }

    /// <summary>
    ///     Removes the element at the specified index from the collection.
    /// </summary>
    /// <param name="index">The index.</param>
    /// <exception cref="System.ArgumentOutOfRangeException">
    ///     Index is less than 0 or index is equal to or greater than Count.
    /// </exception>
    public void RemoveAt(int index)
    {
        _propDefSortOrderPairList.RemoveAt(index);
    }

    /// <summary>
    ///     Tries to get the value for a property definition in the collection.
    /// </summary>
    /// <param name="propertyDefinition">The property definition.</param>
    /// <param name="sortDirection">The sort direction.</param>
    /// <returns>True if collection contains property definition, otherwise false.</returns>
    public bool TryGetValue(PropertyDefinitionBase propertyDefinition, out SortDirection sortDirection)
    {
        foreach (var pair in _propDefSortOrderPairList)
        {
            if (pair.Value.Equals(propertyDefinition))
            {
                sortDirection = pair.Value;
                return true;
            }
        }

        sortDirection = SortDirection.Ascending; // out parameter has to be set to some value.
        return false;
    }

    /// <summary>
    ///     Writes to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="xmlElementName">Name of the XML element.</param>
    internal void WriteToXml(EwsServiceXmlWriter writer, string xmlElementName)
    {
        if (Count > 0)
        {
            writer.WriteStartElement(XmlNamespace.Messages, xmlElementName);

            foreach (var keyValuePair in this)
            {
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.FieldOrder);

                writer.WriteAttributeValue(XmlAttributeNames.Order, keyValuePair.Value);
                keyValuePair.Key.WriteToXml(writer);

                writer.WriteEndElement(); // FieldOrder
            }

            writer.WriteEndElement();
        }
    }
}
