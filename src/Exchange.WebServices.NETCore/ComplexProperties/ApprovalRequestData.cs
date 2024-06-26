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
///     Represents approval request information.
/// </summary>
[PublicAPI]
public sealed class ApprovalRequestData : ComplexProperty
{
    /// <summary>
    ///     Gets a value indicating whether this is an undecided approval request.
    /// </summary>
    public bool IsUndecidedApprovalRequest { get; private set; }

    /// <summary>
    ///     Gets the approval decision on the request.
    /// </summary>
    public int ApprovalDecision { get; private set; }

    /// <summary>
    ///     Gets the name of the user who made the decision.
    /// </summary>
    public string ApprovalDecisionMaker { get; private set; }

    /// <summary>
    ///     Gets the time at which the decision was made.
    /// </summary>
    public DateTime ApprovalDecisionTime { get; private set; }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ApprovalRequestData" /> class.
    /// </summary>
    internal ApprovalRequestData()
    {
    }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>True if element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        switch (reader.LocalName)
        {
            case XmlElementNames.IsUndecidedApprovalRequest:
            {
                IsUndecidedApprovalRequest = reader.ReadElementValue<bool>();
                return true;
            }
            case XmlElementNames.ApprovalDecision:
            {
                ApprovalDecision = reader.ReadElementValue<int>();
                return true;
            }
            case XmlElementNames.ApprovalDecisionMaker:
            {
                ApprovalDecisionMaker = reader.ReadElementValue<string>();
                return true;
            }
            case XmlElementNames.ApprovalDecisionTime:
            {
                ApprovalDecisionTime = reader.ReadElementValueAsDateTime().Value;
                return true;
            }
            default:
            {
                return false;
            }
        }
    }
}
