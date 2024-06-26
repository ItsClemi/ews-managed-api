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
///     Represents retention policy tag object.
/// </summary>
[PublicAPI]
public sealed class RetentionPolicyTag
{
    /// <summary>
    ///     Retention policy tag display name.
    /// </summary>
    public string? DisplayName { get; set; }

    /// <summary>
    ///     Retention Id.
    /// </summary>
    public Guid RetentionId { get; set; }

    /// <summary>
    ///     Retention period in time span.
    /// </summary>
    public int RetentionPeriod { get; set; }

    /// <summary>
    ///     Retention type.
    /// </summary>
    public ElcFolderType Type { get; set; }

    /// <summary>
    ///     Retention action.
    /// </summary>
    public RetentionActionType RetentionAction { get; set; }

    /// <summary>
    ///     Retention policy tag description.
    /// </summary>
    public string? Description { get; set; }

    /// <summary>
    ///     Is this a visible tag?
    /// </summary>
    public bool IsVisible { get; set; }

    /// <summary>
    ///     Is this a opted into tag?
    /// </summary>
    public bool OptedInto { get; set; }

    /// <summary>
    ///     Is this an archive tag?
    /// </summary>
    public bool IsArchive { get; set; }

    /// <summary>
    ///     Constructor
    /// </summary>
    public RetentionPolicyTag()
    {
    }

    /// <summary>
    ///     Constructor for retention policy tag.
    /// </summary>
    /// <param name="displayName">Display name.</param>
    /// <param name="retentionId">Retention id.</param>
    /// <param name="retentionPeriod">Retention period.</param>
    /// <param name="type">Retention folder type.</param>
    /// <param name="retentionAction">Retention action.</param>
    /// <param name="isVisible">Is visible.</param>
    /// <param name="optedInto">Opted into.</param>
    /// <param name="isArchive">Is archive tag.</param>
    internal RetentionPolicyTag(
        string displayName,
        Guid retentionId,
        int retentionPeriod,
        ElcFolderType type,
        RetentionActionType retentionAction,
        bool isVisible,
        bool optedInto,
        bool isArchive
    )
    {
        DisplayName = displayName;
        RetentionId = retentionId;
        RetentionPeriod = retentionPeriod;
        Type = type;
        RetentionAction = retentionAction;
        IsVisible = isVisible;
        OptedInto = optedInto;
        IsArchive = isArchive;
    }

    /// <summary>
    ///     Load from xml.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>Retention policy tag object.</returns>
    internal static RetentionPolicyTag LoadFromXml(EwsServiceXmlReader reader)
    {
        reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, XmlElementNames.RetentionPolicyTag);

        var retentionPolicyTag = new RetentionPolicyTag
        {
            DisplayName = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.DisplayName),
            RetentionId = new Guid(reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.RetentionId)),
            RetentionPeriod = reader.ReadElementValue<int>(XmlNamespace.Types, XmlElementNames.RetentionPeriod),
            Type = reader.ReadElementValue<ElcFolderType>(XmlNamespace.Types, XmlElementNames.Type),
            RetentionAction = reader.ReadElementValue<RetentionActionType>(
                XmlNamespace.Types,
                XmlElementNames.RetentionAction
            ),
        };

        // Description is not a required property.
        reader.Read();
        if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Description))
        {
            retentionPolicyTag.Description = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Description);
        }

        retentionPolicyTag.IsVisible = reader.ReadElementValue<bool>(XmlNamespace.Types, XmlElementNames.IsVisible);
        retentionPolicyTag.OptedInto = reader.ReadElementValue<bool>(XmlNamespace.Types, XmlElementNames.OptedInto);
        retentionPolicyTag.IsArchive = reader.ReadElementValue<bool>(XmlNamespace.Types, XmlElementNames.IsArchive);

        return retentionPolicyTag;
    }
}
