// ---------------------------------------------------------------------------
// <copyright file="ComputedInsightValue.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Implements the class for company insight value.</summary>
//-----------------------------------------------------------------------

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents the ComputedInsightValue.
/// </summary>
[PublicAPI]
public sealed class ComputedInsightValue : InsightValue
{
    /// <summary>
    ///     Gets the collection of computed insight
    ///     value properties.
    /// </summary>
    public ComputedInsightValuePropertyCollection Properties { get; internal set; }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">XML reader</param>
    /// <returns>Whether the element was read</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.InsightSource:
            {
                InsightSource = reader.ReadElementValue<string>();
                break;
            }
            case XmlElementNames.Properties:
            {
                Properties = new ComputedInsightValuePropertyCollection();
                Properties.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.Properties);
                break;
            }
            default:
            {
                return base.TryReadElementFromXml(reader);
            }
        }

        return true;
    }
}
