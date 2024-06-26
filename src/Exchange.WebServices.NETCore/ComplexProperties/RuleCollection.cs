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

/// <summary>
///     Represents a collection of rules.
/// </summary>
[PublicAPI]
public sealed class RuleCollection : ComplexProperty, IEnumerable<Rule>
{
    /// <summary>
    ///     The rules in the rule collection.
    /// </summary>
    private readonly List<Rule> _rules;

    /// <summary>
    ///     Gets a value indicating whether an Outlook rule blob exists in the user's
    ///     mailbox. To update rules with EWS when the Outlook rule blob exists, call
    ///     SetInboxRules passing true as the value of the removeOutlookBlob parameter.
    /// </summary>
    public bool OutlookRuleBlobExists { get; internal set; }

    /// <summary>
    ///     Gets the number of rules in this collection.
    /// </summary>
    public int Count => _rules.Count;

    /// <summary>
    ///     Gets the rule at the specified index in the collection.
    /// </summary>
    /// <param name="index">The index of the rule to get.</param>
    /// <returns>The rule at the specified index.</returns>
    public Rule this[int index]
    {
        get
        {
            if (index < 0 || index >= _rules.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            return _rules[index];
        }
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="RuleCollection" /> class.
    /// </summary>
    internal RuleCollection()
    {
        _rules = new List<Rule>();
    }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Rule))
        {
            var rule = new Rule();
            rule.LoadFromXml(reader, XmlElementNames.Rule);
            _rules.Add(rule);
            return true;
        }

        return false;
    }


    #region IEnumerable Interface

    /// <summary>
    ///     Get an enumerator for the collection
    /// </summary>
    /// <returns>Enumerator</returns>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    /// <summary>
    ///     Get an enumerator for the collection
    /// </summary>
    /// <returns>Enumerator</returns>
    public IEnumerator<Rule> GetEnumerator()
    {
        return _rules.GetEnumerator();
    }

    #endregion
}
