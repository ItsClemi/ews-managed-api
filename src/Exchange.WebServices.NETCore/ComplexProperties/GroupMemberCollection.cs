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
///     Represents a collection of members of GroupMember type.
/// </summary>
[PublicAPI]
public sealed class GroupMemberCollection : ComplexPropertyCollection<GroupMember>, ICustomUpdateSerializer
{
    /// <summary>
    ///     If the collection is cleared, then store PDL members collection is updated with "SetItemField".
    ///     If the collection is not cleared, then store PDL members collection is updated with "AppendToItemField".
    /// </summary>
    private bool _collectionIsCleared;

    /// <summary>
    ///     Initializes a new instance of the <see cref="GroupMemberCollection" /> class.
    /// </summary>
    public GroupMemberCollection()
    {
    }

    /// <summary>
    ///     Writes the update to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="ownerObject">The ews object.</param>
    /// <param name="propertyDefinition">Property definition.</param>
    /// <returns>True if property generated serialization.</returns>
    bool ICustomUpdateSerializer.WriteSetUpdateToXml(
        EwsServiceXmlWriter writer,
        ServiceObject ownerObject,
        PropertyDefinition propertyDefinition
    )
    {
        if (_collectionIsCleared)
        {
            if (AddedItems.Count == 0)
            {
                // Delete the whole members collection
                WriteDeleteMembersCollectionToXml(writer);
            }
            else
            {
                // The collection is cleared, so Set
                WriteSetOrAppendMembersToXml(writer, AddedItems, true);
            }
        }
        else
        {
            // The collection is not cleared, i.e. dl.Members.Clear() is not called.
            // Append AddedItems.
            WriteSetOrAppendMembersToXml(writer, AddedItems, false);

            // Since member replacement is not supported by server
            // Delete old ModifiedItems, then recreate new instead.
            WriteDeleteMembersToXml(writer, ModifiedItems);
            WriteSetOrAppendMembersToXml(writer, ModifiedItems, false);

            // Delete RemovedItems.
            WriteDeleteMembersToXml(writer, RemovedItems);
        }

        return true;
    }

    /// <summary>
    ///     Writes the deletion update to XML.
    /// </summary>
    /// <param name="writer">The writer.</param>
    /// <param name="ewsObject">The ews object.</param>
    /// <returns>True if property generated serialization.</returns>
    bool ICustomUpdateSerializer.WriteDeleteUpdateToXml(EwsServiceXmlWriter writer, ServiceObject ewsObject)
    {
        return false;
    }

    /// <summary>
    ///     Finds the member with the specified key in the collection.
    ///     Members that have not yet been saved do not have a key.
    /// </summary>
    /// <param name="key">The key of the member to find.</param>
    /// <returns>The member with the specified key.</returns>
    public GroupMember? Find(string key)
    {
        EwsUtilities.ValidateParam(key);

        foreach (var item in Items)
        {
            if (item.Key == key)
            {
                return item;
            }
        }

        return null;
    }

    /// <summary>
    ///     Clears the collection.
    /// </summary>
    public void Clear()
    {
        // mark the whole collection for deletion
        InternalClear();
        _collectionIsCleared = true;
    }

    /// <summary>
    ///     Adds a member to the collection.
    /// </summary>
    /// <param name="member">The member to add.</param>
    public void Add(GroupMember member)
    {
        EwsUtilities.ValidateParam(member);

        EwsUtilities.Assert(member.Key == null, "GroupMemberCollection.Add", "member.Key is not null.");

        EwsUtilities.Assert(!Contains(member), "GroupMemberCollection.Add", "The member is already in the collection");

        InternalAdd(member);
    }

    /// <summary>
    ///     Adds multiple members to the collection.
    /// </summary>
    /// <param name="members">The members to add.</param>
    public void AddRange(IEnumerable<GroupMember> members)
    {
        EwsUtilities.ValidateParam(members);

        foreach (var member in members)
        {
            Add(member);
        }
    }

    /// <summary>
    ///     Adds a member linked to a Contact Group.
    /// </summary>
    /// <param name="contactGroupId">The Id of the contact group.</param>
    public void AddContactGroup(ItemId contactGroupId)
    {
        Add(new GroupMember(contactGroupId));
    }

    /// <summary>
    ///     Adds a member linked to a specific contact's e-mail address.
    /// </summary>
    /// <param name="contactId">The Id of the contact.</param>
    /// <param name="addressToLink">The contact's address to link to.</param>
    public void AddPersonalContact(ItemId contactId, string? addressToLink)
    {
        Add(new GroupMember(contactId, addressToLink));
    }

    /// <summary>
    ///     Adds a member linked to a contact's first available e-mail address.
    /// </summary>
    /// <param name="contactId">The Id of the contact.</param>
    public void AddPersonalContact(ItemId contactId)
    {
        AddPersonalContact(contactId, null);
    }

    /// <summary>
    ///     Adds a member linked to an Active Directory user.
    /// </summary>
    /// <param name="smtpAddress">The SMTP address of the member.</param>
    public void AddDirectoryUser(string smtpAddress)
    {
        AddDirectoryUser(smtpAddress, EmailAddress.SmtpRoutingType);
    }

    /// <summary>
    ///     Adds a member linked to an Active Directory user.
    /// </summary>
    /// <param name="address">The address of the member.</param>
    /// <param name="routingType">The routing type of the address.</param>
    public void AddDirectoryUser(string address, string routingType)
    {
        Add(new GroupMember(address, routingType, MailboxType.Mailbox));
    }

    /// <summary>
    ///     Adds a member linked to an Active Directory contact.
    /// </summary>
    /// <param name="smtpAddress">The SMTP address of the Active Directory contact.</param>
    public void AddDirectoryContact(string smtpAddress)
    {
        AddDirectoryContact(smtpAddress, EmailAddress.SmtpRoutingType);
    }

    /// <summary>
    ///     Adds a member linked to an Active Directory contact.
    /// </summary>
    /// <param name="address">The address of the Active Directory contact.</param>
    /// <param name="routingType">The routing type of the address.</param>
    public void AddDirectoryContact(string address, string routingType)
    {
        Add(new GroupMember(address, routingType, MailboxType.Contact));
    }

    /// <summary>
    ///     Adds a member linked to a Public Group.
    /// </summary>
    /// <param name="smtpAddress">The SMTP address of the Public Group.</param>
    public void AddPublicGroup(string smtpAddress)
    {
        Add(new GroupMember(smtpAddress, EmailAddress.SmtpRoutingType, MailboxType.PublicGroup));
    }

    /// <summary>
    ///     Adds a member linked to a mail-enabled Public Folder.
    /// </summary>
    /// <param name="smtpAddress">The SMTP address of the mail-enabled Public Folder.</param>
    public void AddDirectoryPublicFolder(string smtpAddress)
    {
        Add(new GroupMember(smtpAddress, EmailAddress.SmtpRoutingType, MailboxType.PublicFolder));
    }

    /// <summary>
    ///     Adds a one-off member.
    /// </summary>
    /// <param name="displayName">The display name of the member.</param>
    /// <param name="address">The address of the member.</param>
    /// <param name="routingType">The routing type of the address.</param>
    public void AddOneOff(string displayName, string address, string routingType)
    {
        Add(new GroupMember(displayName, address, routingType));
    }

    /// <summary>
    ///     Adds a one-off member.
    /// </summary>
    /// <param name="displayName">The display name of the member.</param>
    /// <param name="smtpAddress">The SMTP address of the member.</param>
    public void AddOneOff(string displayName, string smtpAddress)
    {
        AddOneOff(displayName, smtpAddress, EmailAddress.SmtpRoutingType);
    }

    /// <summary>
    ///     Adds a member that is linked to a specific e-mail address of a contact.
    /// </summary>
    /// <param name="contact">The contact to link to.</param>
    /// <param name="emailAddressKey">The contact's e-mail address to link to.</param>
    public void AddContactEmailAddress(Contact contact, EmailAddressKey emailAddressKey)
    {
        Add(new GroupMember(contact, emailAddressKey));
    }

    /// <summary>
    ///     Removes a member at the specified index.
    /// </summary>
    /// <param name="index">The index of the member to remove.</param>
    public void RemoveAt(int index)
    {
        if (index < 0 || index >= Count)
        {
            throw new ArgumentOutOfRangeException(nameof(index), Strings.IndexIsOutOfRange);
        }

        InternalRemoveAt(index);
    }

    /// <summary>
    ///     Removes a member from the collection.
    /// </summary>
    /// <param name="member">The member to remove.</param>
    /// <returns>True if the group member was successfully removed from the collection, false otherwise.</returns>
    public bool Remove(GroupMember member)
    {
        return InternalRemove(member);
    }

    /// <summary>
    ///     Creates a GroupMember object from an XML element name.
    /// </summary>
    /// <param name="xmlElementName">The XML element name from which to create the e-mail address.</param>
    /// <returns>An GroupMember object.</returns>
    internal override GroupMember CreateComplexProperty(string xmlElementName)
    {
        return new GroupMember();
    }

    /// <summary>
    ///     Clears the change log.
    /// </summary>
    internal override void ClearChangeLog()
    {
        base.ClearChangeLog();
        _collectionIsCleared = false;
    }

    /// <summary>
    ///     Retrieves the XML element name corresponding to the provided GroupMember object.
    /// </summary>
    /// <param name="member">The GroupMember object from which to determine the XML element name.</param>
    /// <returns>The XML element name corresponding to the provided GroupMember object.</returns>
    internal override string GetCollectionItemXmlElementName(GroupMember member)
    {
        return XmlElementNames.Member;
    }

    /// <summary>
    ///     Delete the whole members collection.
    /// </summary>
    /// <param name="writer">Xml writer.</param>
    private static void WriteDeleteMembersCollectionToXml(EwsServiceXmlWriter writer)
    {
        writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.DeleteItemField);
        ContactGroupSchema.Members.WriteToXml(writer);
        writer.WriteEndElement();
    }

    /// <summary>
    ///     Generate XML to delete individual members.
    /// </summary>
    /// <param name="writer">Xml writer.</param>
    /// <param name="members">Members to delete.</param>
    private static void WriteDeleteMembersToXml(EwsServiceXmlWriter writer, List<GroupMember> members)
    {
        if (members.Count != 0)
        {
            var memberPropDef = new GroupMemberPropertyDefinition();

            foreach (var member in members)
            {
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.DeleteItemField);

                memberPropDef.Key = member.Key;
                memberPropDef.WriteToXml(writer);

                writer.WriteEndElement(); // DeleteItemField
            }
        }
    }

    /// <summary>
    ///     Generate XML to Set or Append members.
    ///     When members are set, the existing PDL member collection is cleared.
    ///     On append members are added to the PDL existing members collection.
    /// </summary>
    /// <param name="writer">Xml writer.</param>
    /// <param name="members">Members to set or append.</param>
    /// <param name="setMode">True - set members, false - append members.</param>
    private static void WriteSetOrAppendMembersToXml(
        EwsServiceXmlWriter writer,
        List<GroupMember> members,
        bool setMode
    )
    {
        if (members.Count != 0)
        {
            writer.WriteStartElement(
                XmlNamespace.Types,
                setMode ? XmlElementNames.SetItemField : XmlElementNames.AppendToItemField
            );

            ContactGroupSchema.Members.WriteToXml(writer);

            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.DistributionList);
            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Members);

            foreach (var member in members)
            {
                member.WriteToXml(writer, XmlElementNames.Member);
            }

            writer.WriteEndElement(); // Members
            writer.WriteEndElement(); // Group
            writer.WriteEndElement(); // setMode ? SetItemField : AppendItemField
        }
    }

    /// <summary>
    ///     Validates this instance.
    /// </summary>
    internal override void InternalValidate()
    {
        base.InternalValidate();

        foreach (var groupMember in ModifiedItems)
        {
            if (string.IsNullOrEmpty(groupMember.Key))
            {
                throw new ServiceValidationException(Strings.ContactGroupMemberCannotBeUpdatedWithoutBeingLoadedFirst);
            }
        }
    }
}
