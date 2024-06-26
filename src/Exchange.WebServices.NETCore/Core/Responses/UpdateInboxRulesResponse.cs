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

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents the response to a UpdateInboxRulesResponse operation.
/// </summary>
internal sealed class UpdateInboxRulesResponse : ServiceResponse
{
    /// <summary>
    ///     Gets the rule operation errors in the response.
    /// </summary>
    internal RuleOperationErrorCollection Errors { get; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="UpdateInboxRulesResponse" /> class.
    /// </summary>
    internal UpdateInboxRulesResponse()
    {
        Errors = new RuleOperationErrorCollection();
    }

    /// <summary>
    ///     Loads extra error details from XML
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <param name="xmlElementName">The current element name of the extra error details.</param>
    /// <returns>
    ///     True if the expected extra details is loaded;
    ///     False if the element name does not match the expected element.
    /// </returns>
    internal override bool LoadExtraErrorDetailsFromXml(EwsServiceXmlReader reader, string xmlElementName)
    {
        if (xmlElementName.Equals(XmlElementNames.MessageXml))
        {
            return base.LoadExtraErrorDetailsFromXml(reader, xmlElementName);
        }

        if (xmlElementName.Equals(XmlElementNames.RuleOperationErrors))
        {
            Errors.LoadFromXml(reader, XmlNamespace.Messages, xmlElementName);
            return true;
        }

        return false;
    }
}
