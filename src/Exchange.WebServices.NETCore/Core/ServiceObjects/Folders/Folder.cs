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
///     Represents a generic folder.
/// </summary>
[PublicAPI]
[ServiceObjectDefinition(XmlElementNames.Folder)]
public class Folder : ServiceObject
{
    /// <summary>
    ///     Initializes an unsaved local instance of <see cref="Folder" />. To bind to an existing folder, use Folder.Bind()
    ///     instead.
    /// </summary>
    /// <param name="service">EWS service to which this object belongs.</param>
    public Folder(ExchangeService service)
        : base(service)
    {
    }

    /// <summary>
    ///     Binds to an existing folder, whatever its actual type is, and loads the specified set of properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to use to bind to the folder.</param>
    /// <param name="id">The Id of the folder to bind to.</param>
    /// <param name="propertySet">The set of properties to load.</param>
    /// <param name="token"></param>
    /// <returns>A Folder instance representing the folder corresponding to the specified Id.</returns>
    public static Task<Folder> Bind(
        ExchangeService service,
        FolderId id,
        PropertySet propertySet,
        CancellationToken token = default
    )
    {
        return service.BindToFolder<Folder>(id, propertySet, token);
    }

    /// <summary>
    ///     Binds to an existing folder, whatever its actual type is, and loads its first class properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to use to bind to the folder.</param>
    /// <param name="id">The Id of the folder to bind to.</param>
    /// <param name="token"></param>
    /// <returns>A Folder instance representing the folder corresponding to the specified Id.</returns>
    public static Task<Folder> Bind(ExchangeService service, FolderId id, CancellationToken token = default)
    {
        return Bind(service, id, PropertySet.FirstClassProperties, token);
    }

    /// <summary>
    ///     Binds to an existing folder, whatever its actual type is, and loads the specified set of properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to use to bind to the folder.</param>
    /// <param name="name">The name of the folder to bind to.</param>
    /// <param name="propertySet">The set of properties to load.</param>
    /// <param name="token"></param>
    /// <returns>A Folder instance representing the folder with the specified name.</returns>
    public static Task<Folder> Bind(
        ExchangeService service,
        WellKnownFolderName name,
        PropertySet propertySet,
        CancellationToken token = default
    )
    {
        return Bind(service, new FolderId(name), propertySet, token);
    }

    /// <summary>
    ///     Binds to an existing folder, whatever its actual type is, and loads its first class properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to use to bind to the folder.</param>
    /// <param name="name">The name of the folder to bind to.</param>
    /// <param name="token"></param>
    /// <returns>A Folder instance representing the folder with the specified name.</returns>
    public static Task<Folder> Bind(
        ExchangeService service,
        WellKnownFolderName name,
        CancellationToken token = default
    )
    {
        return Bind(service, new FolderId(name), PropertySet.FirstClassProperties, token);
    }

    /// <summary>
    ///     Validates this instance.
    /// </summary>
    internal override void Validate()
    {
        base.Validate();

        // Validate folder permissions
        if (PropertyBag.Contains(FolderSchema.Permissions))
        {
            Permissions.Validate();
        }
    }

    /// <summary>
    ///     Internal method to return the schema associated with this type of object.
    /// </summary>
    /// <returns>The schema associated with this type of object.</returns>
    internal override ServiceObjectSchema GetSchema()
    {
        return FolderSchema.Instance;
    }

    /// <summary>
    ///     Gets the minimum required server version.
    /// </summary>
    /// <returns>Earliest Exchange version in which this service object type is supported.</returns>
    internal override ExchangeVersion GetMinimumRequiredServerVersion()
    {
        return ExchangeVersion.Exchange2007_SP1;
    }

    /// <summary>
    ///     Gets the name of the change XML element.
    /// </summary>
    /// <returns>XML element name,</returns>
    internal override string GetChangeXmlElementName()
    {
        return XmlElementNames.FolderChange;
    }

    /// <summary>
    ///     Gets the name of the set field XML element.
    /// </summary>
    /// <returns>XML element name,</returns>
    internal override string GetSetFieldXmlElementName()
    {
        return XmlElementNames.SetFolderField;
    }

    /// <summary>
    ///     Gets the name of the delete field XML element.
    /// </summary>
    /// <returns>XML element name,</returns>
    internal override string GetDeleteFieldXmlElementName()
    {
        return XmlElementNames.DeleteFolderField;
    }

    /// <summary>
    ///     Loads the specified set of properties on the object.
    /// </summary>
    /// <param name="propertySet">The properties to load.</param>
    /// <param name="token"></param>
    internal override Task<ServiceResponseCollection<ServiceResponse>> InternalLoad(
        PropertySet propertySet,
        CancellationToken token
    )
    {
        ThrowIfThisIsNew();

        return Service.LoadPropertiesForFolder(this, propertySet, token);
    }

    /// <summary>
    ///     Deletes the object.
    /// </summary>
    /// <param name="deleteMode">The deletion mode.</param>
    /// <param name="sendCancellationsMode">Indicates whether meeting cancellation messages should be sent.</param>
    /// <param name="affectedTaskOccurrences">Indicate which occurrence of a recurring task should be deleted.</param>
    /// <param name="token"></param>
    internal override Task<ServiceResponseCollection<ServiceResponse>> InternalDelete(
        DeleteMode deleteMode,
        SendCancellationsMode? sendCancellationsMode,
        AffectedTaskOccurrence? affectedTaskOccurrences,
        CancellationToken token
    )
    {
        ThrowIfThisIsNew();

        return Service.DeleteFolder(Id, deleteMode, token);
    }

    /// <summary>
    ///     Deletes the folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="deleteMode">Deletion mode.</param>
    /// <param name="token"></param>
    public Task<ServiceResponseCollection<ServiceResponse>> Delete(
        DeleteMode deleteMode,
        CancellationToken token = default
    )
    {
        return InternalDelete(deleteMode, null, null, token);
    }

    /// <summary>
    ///     Empties the folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="deleteMode">The deletion mode.</param>
    /// <param name="deleteSubFolders">Indicates whether sub-folders should also be deleted.</param>
    /// <param name="token"></param>
    public Task<ServiceResponseCollection<ServiceResponse>> Empty(
        DeleteMode deleteMode,
        bool deleteSubFolders,
        CancellationToken token = default
    )
    {
        ThrowIfThisIsNew();
        return Service.EmptyFolder(Id, deleteMode, deleteSubFolders, token);
    }

    /// <summary>
    ///     Marks all items in folder as read. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="suppressReadReceipts">If true, suppress sending read receipts for items.</param>
    /// <param name="token"></param>
    public Task<ServiceResponseCollection<ServiceResponse>> MarkAllItemsAsRead(
        bool suppressReadReceipts,
        CancellationToken token = default
    )
    {
        ThrowIfThisIsNew();
        return Service.MarkAllItemsAsRead(Id, true, suppressReadReceipts, token);
    }

    /// <summary>
    ///     Marks all items in folder as read. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="suppressReadReceipts">If true, suppress sending read receipts for items.</param>
    /// <param name="token"></param>
    public Task<ServiceResponseCollection<ServiceResponse>> MarkAllItemsAsUnread(
        bool suppressReadReceipts,
        CancellationToken token = default
    )
    {
        ThrowIfThisIsNew();
        return Service.MarkAllItemsAsRead(Id, false, suppressReadReceipts, token);
    }

    /// <summary>
    ///     Saves this folder in a specific folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="parentFolderId">The Id of the folder in which to save this folder.</param>
    /// <param name="token"></param>
    public async System.Threading.Tasks.Task Save(FolderId parentFolderId, CancellationToken token = default)
    {
        ThrowIfThisIsNotNew();

        EwsUtilities.ValidateParam(parentFolderId);

        if (IsDirty)
        {
            await Service.CreateFolder(this, parentFolderId, token).ConfigureAwait(false);
        }
    }

    /// <summary>
    ///     Saves this folder in a specific folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="parentFolderName">The name of the folder in which to save this folder.</param>
    public System.Threading.Tasks.Task Save(WellKnownFolderName parentFolderName)
    {
        return Save(new FolderId(parentFolderName));
    }

    /// <summary>
    ///     Applies the local changes that have been made to this folder. Calling this method results in a call to EWS.
    /// </summary>
    public async System.Threading.Tasks.Task Update(CancellationToken token = default)
    {
        if (IsDirty)
        {
            if (PropertyBag.GetIsUpdateCallNecessary())
            {
                await Service.UpdateFolder(this, token).ConfigureAwait(false);
            }
        }
    }

    /// <summary>
    ///     Copies this folder into a specific folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="destinationFolderId">The Id of the folder in which to copy this folder.</param>
    /// <param name="token"></param>
    /// <returns>A Folder representing the copy of this folder.</returns>
    public Task<Folder> Copy(FolderId destinationFolderId, CancellationToken token = default)
    {
        ThrowIfThisIsNew();

        EwsUtilities.ValidateParam(destinationFolderId);

        return Service.CopyFolder(Id, destinationFolderId, token);
    }

    /// <summary>
    ///     Copies this folder into the specified folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="destinationFolderName">The name of the folder in which to copy this folder.</param>
    /// <returns>A Folder representing the copy of this folder.</returns>
    public Task<Folder> Copy(WellKnownFolderName destinationFolderName)
    {
        return Copy(new FolderId(destinationFolderName));
    }

    /// <summary>
    ///     Moves this folder to a specific folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="destinationFolderId">The Id of the folder in which to move this folder.</param>
    /// <param name="token"></param>
    /// <returns>
    ///     A new folder representing this folder in its new location. After Move completes, this folder does not exist
    ///     anymore.
    /// </returns>
    public Task<Folder> Move(FolderId destinationFolderId, CancellationToken token = default)
    {
        ThrowIfThisIsNew();

        EwsUtilities.ValidateParam(destinationFolderId);

        return Service.MoveFolder(Id, destinationFolderId, token);
    }

    /// <summary>
    ///     Moves this folder to the specified folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="destinationFolderName">The name of the folder in which to move this folder.</param>
    /// <returns>
    ///     A new folder representing this folder in its new location. After Move completes, this folder does not exist
    ///     anymore.
    /// </returns>
    public Task<Folder> Move(WellKnownFolderName destinationFolderName)
    {
        return Move(new FolderId(destinationFolderName));
    }

    /// <summary>
    ///     Find items.
    /// </summary>
    /// <typeparam name="TItem">The type of the item.</typeparam>
    /// <param name="queryString">query string to be used for indexed search</param>
    /// <param name="view">The view controlling the number of items returned.</param>
    /// <param name="groupBy">The group by.</param>
    /// <param name="token"></param>
    /// <returns>FindItems response collection.</returns>
    internal Task<ServiceResponseCollection<FindItemResponse<TItem>>> InternalFindItems<TItem>(
        string queryString,
        ViewBase view,
        Grouping? groupBy,
        CancellationToken token
    )
        where TItem : Item
    {
        ThrowIfThisIsNew();

        return Service.FindItems<TItem>(
            [Id,],
            null,
            queryString,
            view,
            groupBy,
            ServiceErrorHandling.ThrowOnError,
            token
        );
    }

    /// <summary>
    ///     Find items.
    /// </summary>
    /// <typeparam name="TItem">The type of the item.</typeparam>
    /// <param name="searchFilter">
    ///     The search filter. Available search filter classes
    ///     include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and
    ///     SearchFilter.SearchFilterCollection
    /// </param>
    /// <param name="view">The view controlling the number of items returned.</param>
    /// <param name="groupBy">The group by.</param>
    /// <param name="token"></param>
    /// <returns>FindItems response collection.</returns>
    internal Task<ServiceResponseCollection<FindItemResponse<TItem>>> InternalFindItems<TItem>(
        SearchFilter? searchFilter,
        ViewBase view,
        Grouping? groupBy,
        CancellationToken token
    )
        where TItem : Item
    {
        ThrowIfThisIsNew();

        return Service.FindItems<TItem>(
            [Id,],
            searchFilter,
            null,
            view,
            groupBy,
            ServiceErrorHandling.ThrowOnError,
            token
        );
    }

    /// <summary>
    ///     Obtains a list of items by searching the contents of this folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="searchFilter">
    ///     The search filter. Available search filter classes
    ///     include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and
    ///     SearchFilter.SearchFilterCollection
    /// </param>
    /// <param name="view">The view controlling the number of items returned.</param>
    /// <param name="token"></param>
    /// <returns>An object representing the results of the search operation.</returns>
    public async Task<FindItemsResults<Item>> FindItems(
        SearchFilter searchFilter,
        ItemView view,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParamAllowNull(searchFilter);

        var responses = await InternalFindItems<Item>(searchFilter, view, null, token).ConfigureAwait(false);

        return responses[0].Results;
    }

    /// <summary>
    ///     Obtains a list of items by searching the contents of this folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="queryString">query string to be used for indexed search</param>
    /// <param name="view">The view controlling the number of items returned.</param>
    /// <param name="token"></param>
    /// <returns>An object representing the results of the search operation.</returns>
    public async Task<FindItemsResults<Item>> FindItems(
        string queryString,
        ItemView view,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParamAllowNull(queryString);

        var responses = await InternalFindItems<Item>(queryString, view, null, token).ConfigureAwait(false);

        return responses[0].Results;
    }

    /// <summary>
    ///     Obtains a list of items by searching the contents of this folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="view">The view controlling the number of items returned.</param>
    /// <param name="token"></param>
    /// <returns>An object representing the results of the search operation.</returns>
    public async Task<FindItemsResults<Item>> FindItems(ItemView view, CancellationToken token = default)
    {
        var responses = await InternalFindItems<Item>((SearchFilter?)null, view, null, token).ConfigureAwait(false);

        return responses[0].Results;
    }

    /// <summary>
    ///     Obtains a grouped list of items by searching the contents of this folder. Calling this method results in a call to
    ///     EWS.
    /// </summary>
    /// <param name="searchFilter">
    ///     The search filter. Available search filter classes
    ///     include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and
    ///     SearchFilter.SearchFilterCollection
    /// </param>
    /// <param name="view">The view controlling the number of items returned.</param>
    /// <param name="groupBy">The grouping criteria.</param>
    /// <param name="token"></param>
    /// <returns>A collection of grouped items representing the contents of this folder.</returns>
    public async Task<GroupedFindItemsResults<Item>> FindItems(
        SearchFilter? searchFilter,
        ItemView view,
        Grouping groupBy,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(groupBy);
        EwsUtilities.ValidateParamAllowNull(searchFilter);

        var responses = await InternalFindItems<Item>(searchFilter, view, groupBy, token).ConfigureAwait(false);

        return responses[0].GroupedFindResults;
    }

    /// <summary>
    ///     Obtains a grouped list of items by searching the contents of this folder. Calling this method results in a call to
    ///     EWS.
    /// </summary>
    /// <param name="queryString">query string to be used for indexed search</param>
    /// <param name="view">The view controlling the number of items returned.</param>
    /// <param name="groupBy">The grouping criteria.</param>
    /// <param name="token"></param>
    /// <returns>A collection of grouped items representing the contents of this folder.</returns>
    public async Task<GroupedFindItemsResults<Item>> FindItems(
        string queryString,
        ItemView view,
        Grouping groupBy,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(groupBy);

        var responses = await InternalFindItems<Item>(queryString, view, groupBy, token).ConfigureAwait(false);

        return responses[0].GroupedFindResults;
    }

    /// <summary>
    ///     Obtains a list of folders by searching the sub-folders of this folder. Calling this method results in a call to
    ///     EWS.
    /// </summary>
    /// <param name="view">The view controlling the number of folders returned.</param>
    /// <returns>An object representing the results of the search operation.</returns>
    public Task<FindFoldersResults> FindFolders(FolderView view)
    {
        ThrowIfThisIsNew();

        return Service.FindFolders(Id, view);
    }

    /// <summary>
    ///     Obtains a list of folders by searching the sub-folders of this folder. Calling this method results in a call to
    ///     EWS.
    /// </summary>
    /// <param name="searchFilter">
    ///     The search filter. Available search filter classes
    ///     include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and
    ///     SearchFilter.SearchFilterCollection
    /// </param>
    /// <param name="view">The view controlling the number of folders returned.</param>
    /// <returns>An object representing the results of the search operation.</returns>
    public Task<FindFoldersResults> FindFolders(SearchFilter searchFilter, FolderView view)
    {
        ThrowIfThisIsNew();

        return Service.FindFolders(Id, searchFilter, view);
    }

    /// <summary>
    ///     Obtains a grouped list of items by searching the contents of this folder. Calling this method results in a call to
    ///     EWS.
    /// </summary>
    /// <param name="view">The view controlling the number of items returned.</param>
    /// <param name="groupBy">The grouping criteria.</param>
    /// <returns>A collection of grouped items representing the contents of this folder.</returns>
    public Task<GroupedFindItemsResults<Item>> FindItems(ItemView view, Grouping groupBy)
    {
        EwsUtilities.ValidateParam(groupBy);

        return FindItems((SearchFilter?)null, view, groupBy);
    }

    /// <summary>
    ///     Get the property definition for the Id property.
    /// </summary>
    /// <returns>A PropertyDefinition instance.</returns>
    internal override PropertyDefinition GetIdPropertyDefinition()
    {
        return FolderSchema.Id;
    }

    /// <summary>
    ///     Sets the extended property.
    /// </summary>
    /// <param name="extendedPropertyDefinition">The extended property definition.</param>
    /// <param name="value">The value.</param>
    public void SetExtendedProperty(ExtendedPropertyDefinition extendedPropertyDefinition, object value)
    {
        ExtendedProperties.SetExtendedProperty(extendedPropertyDefinition, value);
    }

    /// <summary>
    ///     Removes an extended property.
    /// </summary>
    /// <param name="extendedPropertyDefinition">The extended property definition.</param>
    /// <returns>True if property was removed.</returns>
    public bool RemoveExtendedProperty(ExtendedPropertyDefinition extendedPropertyDefinition)
    {
        return ExtendedProperties.RemoveExtendedProperty(extendedPropertyDefinition);
    }

    /// <summary>
    ///     Gets a list of extended properties defined on this object.
    /// </summary>
    /// <returns>Extended properties collection.</returns>
    internal override ExtendedPropertyCollection? GetExtendedProperties()
    {
        return ExtendedProperties;
    }


    #region Properties

    /// <summary>
    ///     Gets the Id of the folder.
    /// </summary>
    public FolderId Id => (FolderId)PropertyBag[GetIdPropertyDefinition()]!;

    /// <summary>
    ///     Gets the Id of this folder's parent folder.
    /// </summary>
    public FolderId ParentFolderId => (FolderId)PropertyBag[FolderSchema.ParentFolderId]!;

    /// <summary>
    ///     Gets the number of child folders this folder has.
    /// </summary>
    public int ChildFolderCount => (int)PropertyBag[FolderSchema.ChildFolderCount]!;

    /// <summary>
    ///     Gets or sets the display name of the folder.
    /// </summary>
    public string DisplayName
    {
        get => (string)PropertyBag[FolderSchema.DisplayName]!;
        set => PropertyBag[FolderSchema.DisplayName] = value;
    }

    /// <summary>
    ///     Gets or sets the custom class name of this folder.
    /// </summary>
    public string FolderClass
    {
        get => (string)PropertyBag[FolderSchema.FolderClass]!;
        set => PropertyBag[FolderSchema.FolderClass] = value;
    }

    /// <summary>
    ///     Gets the total number of items contained in the folder.
    /// </summary>
    public int TotalCount => (int)PropertyBag[FolderSchema.TotalCount]!;

    /// <summary>
    ///     Gets a list of extended properties associated with the folder.
    /// </summary>
    public ExtendedPropertyCollection ExtendedProperties =>
        (ExtendedPropertyCollection)PropertyBag[ServiceObjectSchema.ExtendedProperties]!;

    /// <summary>
    ///     Gets the Email Lifecycle Management (ELC) information associated with the folder.
    /// </summary>
    public ManagedFolderInformation ManagedFolderInformation =>
        (ManagedFolderInformation)PropertyBag[FolderSchema.ManagedFolderInformation]!;

    /// <summary>
    ///     Gets a value indicating the effective rights the current authenticated user has on the folder.
    /// </summary>
    public EffectiveRights EffectiveRights => (EffectiveRights)PropertyBag[FolderSchema.EffectiveRights]!;

    /// <summary>
    ///     Gets a list of permissions for the folder.
    /// </summary>
    public FolderPermissionCollection Permissions => (FolderPermissionCollection)PropertyBag[FolderSchema.Permissions]!;

    /// <summary>
    ///     Gets the number of unread items in the folder.
    /// </summary>
    public int UnreadCount => (int)PropertyBag[FolderSchema.UnreadCount]!;

    /// <summary>
    ///     Gets or sets the policy tag.
    /// </summary>
    public PolicyTag PolicyTag
    {
        get => (PolicyTag)PropertyBag[FolderSchema.PolicyTag]!;
        set => PropertyBag[FolderSchema.PolicyTag] = value;
    }

    /// <summary>
    ///     Gets or sets the archive tag.
    /// </summary>
    public ArchiveTag ArchiveTag
    {
        get => (ArchiveTag)PropertyBag[FolderSchema.ArchiveTag]!;
        set => PropertyBag[FolderSchema.ArchiveTag] = value;
    }

    /// <summary>
    ///     Gets the well known name of this folder, if any, as a string.
    /// </summary>
    /// <value>The well known name of this folder as a string, or null if this folder isn't a well known folder.</value>
    public string WellKnownFolderNameAsString => (string)PropertyBag[FolderSchema.WellKnownFolderName]!;

    /// <summary>
    ///     Gets the well known name of this folder, if any.
    /// </summary>
    /// <value>The well known name of this folder, or null if this folder isn't a well known folder.</value>
    public WellKnownFolderName? WellKnownFolderName
    {
        get
        {
            if (EwsUtilities.TryParse(WellKnownFolderNameAsString, out WellKnownFolderName result))
            {
                return result;
            }

            return null;
        }
    }

    #endregion
}
