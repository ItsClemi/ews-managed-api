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
///     Represents the permissions of a delegate user.
/// </summary>
[PublicAPI]
public sealed class DelegatePermissions : ComplexProperty
{
    private readonly Dictionary<string, DelegateFolderPermission> _delegateFolderPermissions;

    /// <summary>
    ///     Gets or sets the delegate user's permission on the principal's calendar.
    /// </summary>
    public DelegateFolderPermissionLevel CalendarFolderPermissionLevel
    {
        get => _delegateFolderPermissions[XmlElementNames.CalendarFolderPermissionLevel].PermissionLevel;
        set => _delegateFolderPermissions[XmlElementNames.CalendarFolderPermissionLevel].PermissionLevel = value;
    }

    /// <summary>
    ///     Gets or sets the delegate user's permission on the principal's tasks folder.
    /// </summary>
    public DelegateFolderPermissionLevel TasksFolderPermissionLevel
    {
        get => _delegateFolderPermissions[XmlElementNames.TasksFolderPermissionLevel].PermissionLevel;
        set => _delegateFolderPermissions[XmlElementNames.TasksFolderPermissionLevel].PermissionLevel = value;
    }

    /// <summary>
    ///     Gets or sets the delegate user's permission on the principal's inbox.
    /// </summary>
    public DelegateFolderPermissionLevel InboxFolderPermissionLevel
    {
        get => _delegateFolderPermissions[XmlElementNames.InboxFolderPermissionLevel].PermissionLevel;
        set => _delegateFolderPermissions[XmlElementNames.InboxFolderPermissionLevel].PermissionLevel = value;
    }

    /// <summary>
    ///     Gets or sets the delegate user's permission on the principal's contacts folder.
    /// </summary>
    public DelegateFolderPermissionLevel ContactsFolderPermissionLevel
    {
        get => _delegateFolderPermissions[XmlElementNames.ContactsFolderPermissionLevel].PermissionLevel;
        set => _delegateFolderPermissions[XmlElementNames.ContactsFolderPermissionLevel].PermissionLevel = value;
    }

    /// <summary>
    ///     Gets or sets the delegate user's permission on the principal's notes folder.
    /// </summary>
    public DelegateFolderPermissionLevel NotesFolderPermissionLevel
    {
        get => _delegateFolderPermissions[XmlElementNames.NotesFolderPermissionLevel].PermissionLevel;
        set => _delegateFolderPermissions[XmlElementNames.NotesFolderPermissionLevel].PermissionLevel = value;
    }

    /// <summary>
    ///     Gets or sets the delegate user's permission on the principal's journal folder.
    /// </summary>
    public DelegateFolderPermissionLevel JournalFolderPermissionLevel
    {
        get => _delegateFolderPermissions[XmlElementNames.JournalFolderPermissionLevel].PermissionLevel;
        set => _delegateFolderPermissions[XmlElementNames.JournalFolderPermissionLevel].PermissionLevel = value;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="DelegatePermissions" /> class.
    /// </summary>
    internal DelegatePermissions()
    {
        _delegateFolderPermissions = new Dictionary<string, DelegateFolderPermission>
        {
            // @formatter:off
            { XmlElementNames.CalendarFolderPermissionLevel, new DelegateFolderPermission() },
            { XmlElementNames.TasksFolderPermissionLevel, new DelegateFolderPermission() },
            { XmlElementNames.InboxFolderPermissionLevel, new DelegateFolderPermission() },
            { XmlElementNames.ContactsFolderPermissionLevel, new DelegateFolderPermission() },
            { XmlElementNames.NotesFolderPermissionLevel, new DelegateFolderPermission() },
            { XmlElementNames.JournalFolderPermissionLevel, new DelegateFolderPermission() },
            // @formatter:on
        };
    }

    /// <summary>
    ///     Resets this instance.
    /// </summary>
    internal void Reset()
    {
        foreach (var delegateFolderPermission in _delegateFolderPermissions.Values)
        {
            delegateFolderPermission.Reset();
        }
    }

    /// <summary>
    ///     Tries to read element from XML.
    /// </summary>
    /// <param name="reader">The reader.</param>
    /// <returns>Returns true if element was read.</returns>
    internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
    {
        if (_delegateFolderPermissions.TryGetValue(reader.LocalName, out var delegateFolderPermission))
        {
            delegateFolderPermission.Initialize(reader.ReadElementValue<DelegateFolderPermissionLevel>());
        }

        return delegateFolderPermission != null;
    }

    /// <summary>
    ///     Writes elements to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
    {
        WritePermissionToXml(writer, XmlElementNames.CalendarFolderPermissionLevel);

        WritePermissionToXml(writer, XmlElementNames.TasksFolderPermissionLevel);

        WritePermissionToXml(writer, XmlElementNames.InboxFolderPermissionLevel);

        WritePermissionToXml(writer, XmlElementNames.ContactsFolderPermissionLevel);

        WritePermissionToXml(writer, XmlElementNames.NotesFolderPermissionLevel);

        WritePermissionToXml(writer, XmlElementNames.JournalFolderPermissionLevel);
    }

    /// <summary>
    ///     Write permission to Xml.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="xmlElementName">The element name.</param>
    private void WritePermissionToXml(EwsServiceXmlWriter writer, string xmlElementName)
    {
        var delegateFolderPermissionLevel = _delegateFolderPermissions[xmlElementName].PermissionLevel;

        // UpdateDelegate fails if Custom permission level is round tripped
        //
        if (delegateFolderPermissionLevel != DelegateFolderPermissionLevel.Custom)
        {
            writer.WriteElementValue(XmlNamespace.Types, xmlElementName, delegateFolderPermissionLevel);
        }
    }

    /// <summary>
    ///     Validates this instance for AddDelegate.
    /// </summary>
    internal void ValidateAddDelegate()
    {
        // If any folder permission is Custom, throw
        //
        if (_delegateFolderPermissions.Any(kvp => kvp.Value.PermissionLevel == DelegateFolderPermissionLevel.Custom))
        {
            throw new ServiceValidationException(Strings.CannotSetDelegateFolderPermissionLevelToCustom);
        }
    }

    /// <summary>
    ///     Validates this instance for UpdateDelegate.
    /// </summary>
    internal void ValidateUpdateDelegate()
    {
        // If any folder permission was changed to custom, throw
        //
        if (_delegateFolderPermissions.Any(
                kvp => kvp.Value.PermissionLevel == DelegateFolderPermissionLevel.Custom &&
                       !kvp.Value.IsExistingPermissionLevelCustom
            ))
        {
            throw new ServiceValidationException(Strings.CannotSetDelegateFolderPermissionLevelToCustom);
        }
    }

    /// <summary>
    ///     Represents a folder's DelegateFolderPermissionLevel
    /// </summary>
    private class DelegateFolderPermission
    {
        /// <summary>
        ///     Gets or sets the delegate user's permission on a principal's folder.
        /// </summary>
        internal DelegateFolderPermissionLevel PermissionLevel { get; set; }

        /// <summary>
        ///     Gets IsExistingPermissionLevelCustom.
        /// </summary>
        internal bool IsExistingPermissionLevelCustom { get; private set; }

        /// <summary>
        ///     Intializes this DelegateFolderPermission.
        /// </summary>
        /// <param name="permissionLevel">The DelegateFolderPermissionLevel</param>
        internal void Initialize(DelegateFolderPermissionLevel permissionLevel)
        {
            PermissionLevel = permissionLevel;
            IsExistingPermissionLevelCustom = permissionLevel == DelegateFolderPermissionLevel.Custom;
        }

        /// <summary>
        ///     Resets this DelegateFolderPermission.
        /// </summary>
        internal void Reset()
        {
            Initialize(DelegateFolderPermissionLevel.None);
        }
    }
}
