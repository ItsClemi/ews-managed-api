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

using System.Collections.ObjectModel;
using System.Globalization;
using System.Xml;

using JetBrains.Annotations;

using Microsoft.Exchange.WebServices.Data.Enumerations;
using Microsoft.Exchange.WebServices.Data.Groups;

// ReSharper disable PossibleMultipleEnumeration

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a binding to the Exchange Web Services.
/// </summary>
[PublicAPI]
public sealed class ExchangeService : ExchangeServiceBase
{
    private const string TargetServerVersionHeaderName = "X-EWS-TargetVersion";

    private string _targetServerVersion;


    private UnifiedMessaging? _unifiedMessaging;


    /// <summary>
    ///     Gets or sets the URL of the Exchange Web Services.
    /// </summary>
    public required Uri Url { get; set; }

    /// <summary>
    ///     Gets or sets the Id of the user that EWS should impersonate.
    /// </summary>
    public ImpersonatedUserId? ImpersonatedUserId { get; set; }

    /// <summary>
    ///     Gets or sets the Id of the user that EWS should open his/her mailbox with privileged logon type.
    /// </summary>
    internal PrivilegedUserId? PrivilegedUserId { get; set; }

    /// <summary>
    /// </summary>
    public ManagementRoles? ManagementRoles { get; set; }

    /// <summary>
    ///     Gets or sets the preferred culture for messages returned by the Exchange Web Services.
    /// </summary>
    public CultureInfo? PreferredCulture { get; set; }

    /// <summary>
    ///     Gets or sets the DateTime precision for DateTime values returned from Exchange Web Services.
    /// </summary>
    public DateTimePrecision DateTimePrecision { get; set; } = DateTimePrecision.Default;

    /// <summary>
    ///     Gets or sets a file attachment content handler.
    /// </summary>
    public IFileAttachmentContentHandler? FileAttachmentContentHandler { get; set; }

    /// <summary>
    ///     Gets the time zone this service is scoped to.
    /// </summary>
    public new TimeZoneInfo TimeZone => base.TimeZone;

    /// <summary>
    ///     Provides access to the Unified Messaging functionalities.
    /// </summary>
    public UnifiedMessaging UnifiedMessaging => _unifiedMessaging ??= new UnifiedMessaging(this);

    /// <summary>
    ///     Gets or sets a value indicating whether the AutodiscoverUrl method should perform SCP (Service Connection Point)
    ///     record lookup when determining
    ///     the Autodiscover service URL.
    /// </summary>
    public bool EnableScpLookup { get; set; } = true;

    /// <summary>
    ///     Gets or sets a value indicating whether Exchange2007 compatibility mode is enabled. (Off by default)
    /// </summary>
    /// <remarks>
    ///     In order to support E12 servers, the Exchange2007CompatibilityMode property can be used
    ///     to indicate that we should use "Exchange2007" as the server version string rather than
    ///     Exchange2007_SP1.
    /// </remarks>
    internal bool Exchange2007CompatibilityMode { get; set; }

    /// <summary>
    ///     Gets or sets a value indicating whether trace output is pretty printed.
    /// </summary>
    [Obsolete]
    public bool TraceEnablePrettyPrinting { get; set; } = true;

    /// <summary>
    ///     Gets or sets the target server version string (newer than Exchange2013).
    /// </summary>
    internal string TargetServerVersion
    {
        get => _targetServerVersion;

        set
        {
            ValidateTargetVersion(value);
            _targetServerVersion = value;
        }
    }


    /// <summary>
    ///     Initializes a new instance of the <see cref="ExchangeService" /> class, targeting
    ///     the latest supported version of EWS and scoped to the system's current time zone.
    /// </summary>
    public ExchangeService()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExchangeService" /> class, targeting
    ///     the latest supported version of EWS and scoped to the specified time zone.
    /// </summary>
    /// <param name="timeZone">The time zone to which the service is scoped.</param>
    public ExchangeService(TimeZoneInfo timeZone)
        : base(timeZone)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExchangeService" /> class, targeting
    ///     the specified version of EWS and scoped to the system's current time zone.
    /// </summary>
    /// <param name="requestedServerVersion">The version of EWS that the service targets.</param>
    public ExchangeService(ExchangeVersion requestedServerVersion)
        : base(requestedServerVersion)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExchangeService" /> class, targeting
    ///     the specified version of EWS and scoped to the specified time zone.
    /// </summary>
    /// <param name="requestedServerVersion">The version of EWS that the service targets.</param>
    /// <param name="timeZone">The time zone to which the service is scoped.</param>
    public ExchangeService(ExchangeVersion requestedServerVersion, TimeZoneInfo timeZone)
        : base(requestedServerVersion, timeZone)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExchangeService" /> class, targeting
    ///     the specified version of EWS and scoped to the system's current time zone.
    /// </summary>
    /// <param name="targetServerVersion">The version (new style) of EWS that the service targets.</param>
    /// <remarks>
    ///     The target version string has a required part and an optional part.
    ///     The required part is two integers separated by a dot, major.minor
    ///     The optional part is a minimum required version, minimum=major.minor
    ///     Examples:
    ///     X-EWS-TargetVersion: 2.4
    ///     X-EWS_TargetVersion: 2.9; minimum=2.4
    /// </remarks>
    internal ExchangeService(string targetServerVersion)
        : base(ExchangeVersion.Exchange2013)
    {
        ValidateTargetVersion(targetServerVersion);
        TargetServerVersion = targetServerVersion;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExchangeService" /> class, targeting
    ///     the specified version of EWS and scoped to the specified time zone.
    /// </summary>
    /// <param name="targetServerVersion">The version (new style) of EWS that the service targets.</param>
    /// <param name="timeZone">The time zone to which the service is scoped.</param>
    /// <remarks>
    ///     The new style version string has a required part and an optional part.
    ///     The required part is two integers separated by a dot, major.minor
    ///     The optional part is a minimum required version, minimum=major.minor
    ///     Examples:
    ///     2.4
    ///     2.9; minimum=2.4
    /// </remarks>
    internal ExchangeService(string targetServerVersion, TimeZoneInfo timeZone)
        : base(ExchangeVersion.Exchange2013, timeZone)
    {
        ValidateTargetVersion(targetServerVersion);
        TargetServerVersion = targetServerVersion;
    }


    #region Response object operations

    /// <summary>
    ///     Create response object.
    /// </summary>
    /// <param name="responseObject">The response object.</param>
    /// <param name="parentFolderId">The parent folder id.</param>
    /// <param name="messageDisposition">The message disposition.</param>
    /// <param name="token"></param>
    /// <returns>The list of items created or modified as a result of the "creation" of the response object.</returns>
    internal async Task<List<Item>> InternalCreateResponseObject(
        ServiceObject responseObject,
        FolderId? parentFolderId,
        MessageDisposition? messageDisposition,
        CancellationToken token
    )
    {
        var request = new CreateResponseObjectRequest(this, ServiceErrorHandling.ThrowOnError)
        {
            ParentFolderId = parentFolderId,
            Items = new[]
            {
                responseObject,
            },
            MessageDisposition = messageDisposition,
        };

        var responses = await request.ExecuteAsync(token).ConfigureAwait(false);
        return responses[0].Items;
    }

    #endregion


    #region PeopleInsights operations

    /// <summary>
    ///     This method is for retreiving people insight for given email addresses
    /// </summary>
    /// <param name="emailAddresses">Specified eamiladdresses to retrieve</param>
    /// <param name="token"></param>
    /// <returns>The collection of Person objects containing the insight info</returns>
    public async Task<Collection<Person>> GetPeopleInsights(
        IEnumerable<string> emailAddresses,
        CancellationToken token = default
    )
    {
        var request = new GetPeopleInsightsRequest(this);
        request.EmailAddresses.AddRange(emailAddresses);

        var response = await request.Execute(token).ConfigureAwait(false);
        return response.People;
    }

    #endregion


    #region MailTips operations

    /// <summary>
    /// Gets MailTips for given users. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="sendingAs">E-mail address that a user is trying to send as.</param>
    /// <param name="recipients">Collection of recipients that would receive a copy of the message.</param>
    /// <param name="requested">Mail tips requested from the service.</param>
    /// <param name="token"></param>
    /// <returns>List of GetMailTips results.</returns>
    public Task<GetMailTipsResults> GetMailTips(
        string sendingAs,
        Mailbox[] recipients,
        MailTipsRequested requested,
        CancellationToken token = default
    )
    {
        var request = new GetMailTipsRequest(this)
        {
            SendingAs = sendingAs,
            Recipients = recipients,
            MailTipsRequested = requested,
        };

        return request.Execute(token);
    }

    #endregion


    #region MRM operations

    /// <summary>
    ///     Get user retention policy tags.
    /// </summary>
    /// <returns>Service response object.</returns>
    public Task<GetUserRetentionPolicyTagsResponse> GetUserRetentionPolicyTags(CancellationToken token = default)
    {
        var request = new GetUserRetentionPolicyTagsRequest(this);

        return request.Execute(token);
    }

    #endregion


    #region Diagnostic Method -- Only used by test

    /// <summary>
    ///     Executes the diagnostic method.
    /// </summary>
    /// <param name="verb">The verb.</param>
    /// <param name="parameter">The parameter.</param>
    /// <param name="token"></param>
    /// <returns></returns>
    internal async Task<XmlDocument> ExecuteDiagnosticMethod(string verb, XmlNode parameter, CancellationToken token)
    {
        var request = new ExecuteDiagnosticMethodRequest(this)
        {
            Verb = verb,
            Parameter = parameter,
        };

        var responses = await request.ExecuteAsync(token).ConfigureAwait(false);
        return responses[0].ReturnValue;
    }

    #endregion


    #region Folder operations

    /// <summary>
    ///     Creates a folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="folder">The folder.</param>
    /// <param name="parentFolderId">The parent folder id.</param>
    /// <param name="token"></param>
    internal System.Threading.Tasks.Task CreateFolder(Folder folder, FolderId parentFolderId, CancellationToken token)
    {
        var request = new CreateFolderRequest(this, ServiceErrorHandling.ThrowOnError)
        {
            Folders = new[]
            {
                folder,
            },
            ParentFolderId = parentFolderId,
        };

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Updates a folder.
    /// </summary>
    /// <param name="folder">The folder.</param>
    /// <param name="token"></param>
    internal System.Threading.Tasks.Task UpdateFolder(Folder folder, CancellationToken token)
    {
        var request = new UpdateFolderRequest(this, ServiceErrorHandling.ThrowOnError);

        request.Folders.Add(folder);

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Copies a folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="folderId">The folder id.</param>
    /// <param name="destinationFolderId">The destination folder id.</param>
    /// <param name="token"></param>
    /// <returns>Copy of folder.</returns>
    internal async Task<Folder> CopyFolder(FolderId folderId, FolderId destinationFolderId, CancellationToken token)
    {
        var request = new CopyFolderRequest(this, ServiceErrorHandling.ThrowOnError)
        {
            DestinationFolderId = destinationFolderId,
        };

        request.FolderIds.Add(folderId);

        var responses = await request.ExecuteAsync(token).ConfigureAwait(false);

        return responses[0].Folder;
    }

    /// <summary>
    ///     Move a folder.
    /// </summary>
    /// <param name="folderId">The folder id.</param>
    /// <param name="destinationFolderId">The destination folder id.</param>
    /// <param name="token"></param>
    /// <returns>Moved folder.</returns>
    internal async Task<Folder> MoveFolder(FolderId folderId, FolderId destinationFolderId, CancellationToken token)
    {
        var request = new MoveFolderRequest(this, ServiceErrorHandling.ThrowOnError)
        {
            DestinationFolderId = destinationFolderId,
        };

        request.FolderIds.Add(folderId);

        var responses = await request.ExecuteAsync(token).ConfigureAwait(false);

        return responses[0].Folder;
    }

    /// <summary>
    ///     Finds folders.
    /// </summary>
    /// <param name="parentFolderIds">The parent folder ids.</param>
    /// <param name="searchFilter">
    ///     The search filter. Available search filter classes
    ///     include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and
    ///     SearchFilter.SearchFilterCollection
    /// </param>
    /// <param name="view">The view controlling the number of folders returned.</param>
    /// <param name="errorHandlingMode">Indicates the type of error handling should be done.</param>
    /// <param name="token"></param>
    /// <returns>Collection of service responses.</returns>
    private Task<ServiceResponseCollection<FindFolderResponse>> InternalFindFolders(
        IEnumerable<FolderId> parentFolderIds,
        SearchFilter? searchFilter,
        FolderView view,
        ServiceErrorHandling errorHandlingMode,
        CancellationToken token
    )
    {
        var request = new FindFolderRequest(this, errorHandlingMode)
        {
            SearchFilter = searchFilter,
            View = view,
        };

        request.ParentFolderIds.AddRange(parentFolderIds);

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Obtains a list of folders by searching the sub-folders of the specified folder.
    /// </summary>
    /// <param name="parentFolderId">The Id of the folder in which to search for folders.</param>
    /// <param name="searchFilter">
    ///     The search filter. Available search filter classes
    ///     include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and
    ///     SearchFilter.SearchFilterCollection
    /// </param>
    /// <param name="view">The view controlling the number of folders returned.</param>
    /// <param name="token"></param>
    /// <returns>An object representing the results of the search operation.</returns>
    public async Task<FindFoldersResults> FindFolders(
        FolderId parentFolderId,
        SearchFilter searchFilter,
        FolderView view,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(parentFolderId);
        EwsUtilities.ValidateParam(view);
        EwsUtilities.ValidateParamAllowNull(searchFilter);

        var responses = await InternalFindFolders(
                new[]
                {
                    parentFolderId,
                },
                searchFilter,
                view,
                ServiceErrorHandling.ThrowOnError,
                token
            )
            .ConfigureAwait(false);

        return responses[0].Results;
    }

    /// <summary>
    ///     Obtains a list of folders by searching the sub-folders of each of the specified folders.
    /// </summary>
    /// <param name="parentFolderIds">The Ids of the folders in which to search for folders.</param>
    /// <param name="searchFilter">
    ///     The search filter. Available search filter classes
    ///     include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and
    ///     SearchFilter.SearchFilterCollection
    /// </param>
    /// <param name="view">The view controlling the number of folders returned.</param>
    /// <param name="token"></param>
    /// <returns>An object representing the results of the search operation.</returns>
    public Task<ServiceResponseCollection<FindFolderResponse>> FindFolders(
        IEnumerable<FolderId> parentFolderIds,
        SearchFilter searchFilter,
        FolderView view,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(parentFolderIds);
        EwsUtilities.ValidateParam(view);
        EwsUtilities.ValidateParamAllowNull(searchFilter);

        return InternalFindFolders(parentFolderIds, searchFilter, view, ServiceErrorHandling.ReturnErrors, token);
    }

    /// <summary>
    ///     Obtains a list of folders by searching the sub-folders of the specified folder.
    /// </summary>
    /// <param name="parentFolderId">The Id of the folder in which to search for folders.</param>
    /// <param name="view">The view controlling the number of folders returned.</param>
    /// <param name="token"></param>
    /// <returns>An object representing the results of the search operation.</returns>
    public async Task<FindFoldersResults> FindFolders(
        FolderId parentFolderId,
        FolderView view,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(parentFolderId);
        EwsUtilities.ValidateParam(view);

        var responses = await InternalFindFolders(
                new[]
                {
                    parentFolderId,
                },
                null,
                view,
                ServiceErrorHandling.ThrowOnError,
                token
            )
            .ConfigureAwait(false);

        return responses[0].Results;
    }

    /// <summary>
    ///     Obtains a list of folders by searching the sub-folders of the specified folder.
    /// </summary>
    /// <param name="parentFolderName">The name of the folder in which to search for folders.</param>
    /// <param name="searchFilter">
    ///     The search filter. Available search filter classes
    ///     include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and
    ///     SearchFilter.SearchFilterCollection
    /// </param>
    /// <param name="view">The view controlling the number of folders returned.</param>
    /// <returns>An object representing the results of the search operation.</returns>
    public Task<FindFoldersResults> FindFolders(
        WellKnownFolderName parentFolderName,
        SearchFilter searchFilter,
        FolderView view
    )
    {
        return FindFolders(new FolderId(parentFolderName), searchFilter, view);
    }

    /// <summary>
    ///     Obtains a list of folders by searching the sub-folders of the specified folder.
    /// </summary>
    /// <param name="parentFolderName">The name of the folder in which to search for folders.</param>
    /// <param name="view">The view controlling the number of folders returned.</param>
    /// <returns>An object representing the results of the search operation.</returns>
    public Task<FindFoldersResults> FindFolders(WellKnownFolderName parentFolderName, FolderView view)
    {
        return FindFolders(new FolderId(parentFolderName), view);
    }

    /// <summary>
    ///     Load specified properties for a folder.
    /// </summary>
    /// <param name="folder">The folder.</param>
    /// <param name="propertySet">The property set.</param>
    /// <param name="token"></param>
    internal Task<ServiceResponseCollection<ServiceResponse>> LoadPropertiesForFolder(
        Folder? folder,
        PropertySet propertySet,
        CancellationToken token
    )
    {
        EwsUtilities.ValidateParam(folder);
        EwsUtilities.ValidateParam(propertySet);

        var request = new GetFolderRequestForLoad(this, ServiceErrorHandling.ThrowOnError);

        request.FolderIds.Add(folder);
        request.PropertySet = propertySet;

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Binds to a folder.
    /// </summary>
    /// <param name="folderId">The folder id.</param>
    /// <param name="propertySet">The property set.</param>
    /// <param name="cancellationToken"></param>
    /// <returns>Folder</returns>
    internal async Task<Folder?> BindToFolder(
        FolderId folderId,
        PropertySet propertySet,
        CancellationToken cancellationToken
    )
    {
        EwsUtilities.ValidateParam(folderId);
        EwsUtilities.ValidateParam(propertySet);

        var responses = await InternalBindToFolders(
                [folderId,],
                propertySet,
                ServiceErrorHandling.ThrowOnError,
                cancellationToken
            )
            .ConfigureAwait(false);

        return responses[0].Folder;
    }

    /// <summary>
    ///     Binds to folder.
    /// </summary>
    /// <typeparam name="TFolder">The type of the folder.</typeparam>
    /// <param name="folderId">The folder id.</param>
    /// <param name="propertySet">The property set.</param>
    /// <param name="cancellationToken"></param>
    /// <returns>Folder</returns>
    internal async Task<TFolder> BindToFolder<TFolder>(
        FolderId folderId,
        PropertySet propertySet,
        CancellationToken cancellationToken
    )
        where TFolder : Folder
    {
        var result = await BindToFolder(folderId, propertySet, cancellationToken).ConfigureAwait(false);

        if (result is not TFolder folder)
        {
            throw new ServiceLocalException(
                string.Format(Strings.FolderTypeNotCompatible, result.GetType().Name, typeof(TFolder).Name)
            );
        }

        return folder;
    }

    /// <summary>
    ///     Binds to multiple folders in a single call to EWS.
    /// </summary>
    /// <param name="folderIds">The Ids of the folders to bind to.</param>
    /// <param name="propertySet">The set of properties to load.</param>
    /// <param name="token"></param>
    /// <returns>A ServiceResponseCollection providing results for each of the specified folder Ids.</returns>
    public Task<ServiceResponseCollection<GetFolderResponse>> BindToFolders(
        IEnumerable<FolderId> folderIds,
        PropertySet propertySet,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParamCollection(folderIds);
        EwsUtilities.ValidateParam(propertySet);

        return InternalBindToFolders(folderIds, propertySet, ServiceErrorHandling.ReturnErrors, token);
    }

    /// <summary>
    ///     Binds to multiple folders in a single call to EWS.
    /// </summary>
    /// <param name="folderIds">The Ids of the folders to bind to.</param>
    /// <param name="propertySet">The set of properties to load.</param>
    /// <param name="errorHandling">Type of error handling to perform.</param>
    /// <param name="cancellationToken"></param>
    /// <returns>A ServiceResponseCollection providing results for each of the specified folder Ids.</returns>
    private Task<ServiceResponseCollection<GetFolderResponse>> InternalBindToFolders(
        IEnumerable<FolderId> folderIds,
        PropertySet propertySet,
        ServiceErrorHandling errorHandling,
        CancellationToken cancellationToken
    )
    {
        var request = new GetFolderRequest(this, errorHandling)
        {
            PropertySet = propertySet,
        };

        request.FolderIds.AddRange(folderIds);

        return request.ExecuteAsync(cancellationToken);
    }

    /// <summary>
    ///     Deletes a folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="folderId">The folder id.</param>
    /// <param name="deleteMode">The delete mode.</param>
    /// <param name="token"></param>
    internal Task<ServiceResponseCollection<ServiceResponse>> DeleteFolder(
        FolderId folderId,
        DeleteMode deleteMode,
        CancellationToken token
    )
    {
        EwsUtilities.ValidateParam(folderId);

        var request = new DeleteFolderRequest(this, ServiceErrorHandling.ThrowOnError)
        {
            DeleteMode = deleteMode,
        };

        request.FolderIds.Add(folderId);

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Empties a folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="folderId">The folder id.</param>
    /// <param name="deleteMode">The delete mode.</param>
    /// <param name="deleteSubFolders">if set to <c>true</c> empty folder should also delete sub folders.</param>
    /// <param name="token"></param>
    internal Task<ServiceResponseCollection<ServiceResponse>> EmptyFolder(
        FolderId folderId,
        DeleteMode deleteMode,
        bool deleteSubFolders,
        CancellationToken token
    )
    {
        EwsUtilities.ValidateParam(folderId);

        var request = new EmptyFolderRequest(this, ServiceErrorHandling.ThrowOnError)
        {
            DeleteMode = deleteMode,
            DeleteSubFolders = deleteSubFolders,
        };

        request.FolderIds.Add(folderId);

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Marks all items in folder as read/unread. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="folderId">The folder id.</param>
    /// <param name="readFlag">If true, items marked as read, otherwise unread.</param>
    /// <param name="suppressReadReceipts">If true, suppress read receipts for items.</param>
    /// <param name="token"></param>
    internal Task<ServiceResponseCollection<ServiceResponse>> MarkAllItemsAsRead(
        FolderId folderId,
        bool readFlag,
        bool suppressReadReceipts,
        CancellationToken token
    )
    {
        EwsUtilities.ValidateParam(folderId);
        EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013, "MarkAllItemsAsRead");

        var request = new MarkAllItemsAsReadRequest(this, ServiceErrorHandling.ThrowOnError)
        {
            ReadFlag = readFlag,
            SuppressReadReceipts = suppressReadReceipts,
        };

        request.FolderIds.Add(folderId);

        return request.ExecuteAsync(token);
    }

    #endregion


    #region Item operations

    /// <summary>
    ///     Creates multiple items in a single EWS call. Supported item classes are EmailMessage, Appointment, Contact,
    ///     PostItem, Task and Item.
    ///     CreateItems does not support items that have unsaved attachments.
    /// </summary>
    /// <param name="items">The items to create.</param>
    /// <param name="parentFolderId">
    ///     The Id of the folder in which to place the newly created items. If null, items are created
    ///     in their default folders.
    /// </param>
    /// <param name="messageDisposition">
    ///     Indicates the disposition mode for items of type EmailMessage. Required if items
    ///     contains at least one EmailMessage instance.
    /// </param>
    /// <param name="sendInvitationsMode">
    ///     Indicates if and how invitations should be sent for items of type Appointment.
    ///     Required if items contains at least one Appointment instance.
    /// </param>
    /// <param name="errorHandling">What type of error handling should be performed.</param>
    /// <param name="token"></param>
    /// <returns>A ServiceResponseCollection providing creation results for each of the specified items.</returns>
    private Task<ServiceResponseCollection<ServiceResponse>> InternalCreateItems(
        IEnumerable<Item> items,
        FolderId parentFolderId,
        MessageDisposition? messageDisposition,
        SendInvitationsMode? sendInvitationsMode,
        ServiceErrorHandling errorHandling,
        CancellationToken token
    )
    {
        var request = new CreateItemRequest(this, errorHandling)
        {
            ParentFolderId = parentFolderId,
            Items = items,
            MessageDisposition = messageDisposition,
            SendInvitationsMode = sendInvitationsMode,
        };

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Creates multiple items in a single EWS call. Supported item classes are EmailMessage, Appointment, Contact,
    ///     PostItem, Task and Item.
    ///     CreateItems does not support items that have unsaved attachments.
    /// </summary>
    /// <param name="items">The items to create.</param>
    /// <param name="parentFolderId">
    ///     The Id of the folder in which to place the newly created items. If null, items are created
    ///     in their default folders.
    /// </param>
    /// <param name="messageDisposition">
    ///     Indicates the disposition mode for items of type EmailMessage. Required if items
    ///     contains at least one EmailMessage instance.
    /// </param>
    /// <param name="sendInvitationsMode">
    ///     Indicates if and how invitations should be sent for items of type Appointment.
    ///     Required if items contains at least one Appointment instance.
    /// </param>
    /// <param name="token"></param>
    /// <returns>A ServiceResponseCollection providing creation results for each of the specified items.</returns>
    public Task<ServiceResponseCollection<ServiceResponse>> CreateItems(
        IEnumerable<Item> items,
        FolderId parentFolderId,
        MessageDisposition? messageDisposition,
        SendInvitationsMode? sendInvitationsMode,
        CancellationToken token = default
    )
    {
        // All items have to be new.
        if (!items.TrueForAll(item => item.IsNew))
        {
            throw new ServiceValidationException(Strings.CreateItemsDoesNotHandleExistingItems);
        }

        // Make sure that all items do *not* have unprocessed attachments.
        if (!items.TrueForAll(item => !item.HasUnprocessedAttachmentChanges()))
        {
            throw new ServiceValidationException(Strings.CreateItemsDoesNotAllowAttachments);
        }

        return InternalCreateItems(
            items,
            parentFolderId,
            messageDisposition,
            sendInvitationsMode,
            ServiceErrorHandling.ReturnErrors,
            token
        );
    }

    /// <summary>
    ///     Creates an item. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="item">The item to create.</param>
    /// <param name="parentFolderId">
    ///     The Id of the folder in which to place the newly created item. If null, the item is
    ///     created in its default folders.
    /// </param>
    /// <param name="messageDisposition">
    ///     Indicates the disposition mode for items of type EmailMessage. Required if item is an
    ///     EmailMessage instance.
    /// </param>
    /// <param name="sendInvitationsMode">
    ///     Indicates if and how invitations should be sent for item of type Appointment.
    ///     Required if item is an Appointment instance.
    /// </param>
    /// <param name="token"></param>
    internal System.Threading.Tasks.Task CreateItem(
        Item item,
        FolderId? parentFolderId,
        MessageDisposition? messageDisposition,
        SendInvitationsMode? sendInvitationsMode,
        CancellationToken token
    )
    {
        return InternalCreateItems(
            new[]
            {
                item,
            },
            parentFolderId,
            messageDisposition,
            sendInvitationsMode,
            ServiceErrorHandling.ThrowOnError,
            token
        );
    }

    /// <summary>
    ///     Updates multiple items in a single EWS call. UpdateItems does not support items that have unsaved attachments.
    /// </summary>
    /// <param name="items">The items to update.</param>
    /// <param name="savedItemsDestinationFolderId">
    ///     The folder in which to save sent messages, meeting invitations or
    ///     cancellations. If null, the messages, meeting invitation or cancellations are saved in the Sent Items folder.
    /// </param>
    /// <param name="conflictResolution">The conflict resolution mode.</param>
    /// <param name="messageDisposition">
    ///     Indicates the disposition mode for items of type EmailMessage. Required if items
    ///     contains at least one EmailMessage instance.
    /// </param>
    /// <param name="sendInvitationsOrCancellationsMode">
    ///     Indicates if and how invitations and/or cancellations should be sent
    ///     for items of type Appointment. Required if items contains at least one Appointment instance.
    /// </param>
    /// <param name="errorHandling">What type of error handling should be performed.</param>
    /// <param name="suppressReadReceipt">Whether to suppress read receipts</param>
    /// <param name="token"></param>
    /// <returns>A ServiceResponseCollection providing update results for each of the specified items.</returns>
    private Task<ServiceResponseCollection<UpdateItemResponse>> InternalUpdateItems(
        IEnumerable<Item> items,
        FolderId savedItemsDestinationFolderId,
        ConflictResolutionMode conflictResolution,
        MessageDisposition? messageDisposition,
        SendInvitationsOrCancellationsMode? sendInvitationsOrCancellationsMode,
        ServiceErrorHandling errorHandling,
        bool suppressReadReceipt,
        CancellationToken token
    )
    {
        var request = new UpdateItemRequest(this, errorHandling)
        {
            SavedItemsDestinationFolder = savedItemsDestinationFolderId,
            MessageDisposition = messageDisposition,
            ConflictResolutionMode = conflictResolution,
            SendInvitationsOrCancellationsMode = sendInvitationsOrCancellationsMode,
            SuppressReadReceipts = suppressReadReceipt,
        };

        request.Items.AddRange(items);

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Updates multiple items in a single EWS call. UpdateItems does not support items that have unsaved attachments.
    /// </summary>
    /// <param name="items">The items to update.</param>
    /// <param name="savedItemsDestinationFolderId">
    ///     The folder in which to save sent messages, meeting invitations or
    ///     cancellations. If null, the messages, meeting invitation or cancellations are saved in the Sent Items folder.
    /// </param>
    /// <param name="conflictResolution">The conflict resolution mode.</param>
    /// <param name="messageDisposition">
    ///     Indicates the disposition mode for items of type EmailMessage. Required if items
    ///     contains at least one EmailMessage instance.
    /// </param>
    /// <param name="sendInvitationsOrCancellationsMode">
    ///     Indicates if and how invitations and/or cancellations should be sent
    ///     for items of type Appointment. Required if items contains at least one Appointment instance.
    /// </param>
    /// <returns>A ServiceResponseCollection providing update results for each of the specified items.</returns>
    public Task<ServiceResponseCollection<UpdateItemResponse>> UpdateItems(
        IEnumerable<Item> items,
        FolderId savedItemsDestinationFolderId,
        ConflictResolutionMode conflictResolution,
        MessageDisposition? messageDisposition,
        SendInvitationsOrCancellationsMode? sendInvitationsOrCancellationsMode
    )
    {
        return UpdateItems(
            items,
            savedItemsDestinationFolderId,
            conflictResolution,
            messageDisposition,
            sendInvitationsOrCancellationsMode,
            false
        );
    }

    /// <summary>
    ///     Updates multiple items in a single EWS call. UpdateItems does not support items that have unsaved attachments.
    /// </summary>
    /// <param name="items">The items to update.</param>
    /// <param name="savedItemsDestinationFolderId">
    ///     The folder in which to save sent messages, meeting invitations or
    ///     cancellations. If null, the messages, meeting invitation or cancellations are saved in the Sent Items folder.
    /// </param>
    /// <param name="conflictResolution">The conflict resolution mode.</param>
    /// <param name="messageDisposition">
    ///     Indicates the disposition mode for items of type EmailMessage. Required if items
    ///     contains at least one EmailMessage instance.
    /// </param>
    /// <param name="sendInvitationsOrCancellationsMode">
    ///     Indicates if and how invitations and/or cancellations should be sent
    ///     for items of type Appointment. Required if items contains at least one Appointment instance.
    /// </param>
    /// <param name="suppressReadReceipts">Whether to suppress read receipts</param>
    /// <param name="token"></param>
    /// <returns>A ServiceResponseCollection providing update results for each of the specified items.</returns>
    public Task<ServiceResponseCollection<UpdateItemResponse>> UpdateItems(
        IEnumerable<Item> items,
        FolderId savedItemsDestinationFolderId,
        ConflictResolutionMode conflictResolution,
        MessageDisposition? messageDisposition,
        SendInvitationsOrCancellationsMode? sendInvitationsOrCancellationsMode,
        bool suppressReadReceipts,
        CancellationToken token = default
    )
    {
        // All items have to exist on the server (!new) and modified (dirty)
        if (!items.TrueForAll(item => !item.IsNew && item.IsDirty))
        {
            throw new ServiceValidationException(Strings.UpdateItemsDoesNotSupportNewOrUnchangedItems);
        }

        // Make sure that all items do *not* have unprocessed attachments.
        if (!items.TrueForAll(item => !item.HasUnprocessedAttachmentChanges()))
        {
            throw new ServiceValidationException(Strings.UpdateItemsDoesNotAllowAttachments);
        }

        return InternalUpdateItems(
            items,
            savedItemsDestinationFolderId,
            conflictResolution,
            messageDisposition,
            sendInvitationsOrCancellationsMode,
            ServiceErrorHandling.ReturnErrors,
            suppressReadReceipts,
            token
        );
    }

    /// <summary>
    ///     Updates an item.
    /// </summary>
    /// <param name="item">The item to update.</param>
    /// <param name="savedItemsDestinationFolderId">
    ///     The folder in which to save sent messages, meeting invitations or
    ///     cancellations. If null, the message, meeting invitation or cancellation is saved in the Sent Items folder.
    /// </param>
    /// <param name="conflictResolution">The conflict resolution mode.</param>
    /// <param name="messageDisposition">
    ///     Indicates the disposition mode for an item of type EmailMessage. Required if item is
    ///     an EmailMessage instance.
    /// </param>
    /// <param name="sendInvitationsOrCancellationsMode">
    ///     Indicates if and how invitations and/or cancellations should be sent
    ///     for ian tem of type Appointment. Required if item is an Appointment instance.
    /// </param>
    /// <param name="token"></param>
    /// <returns>Updated item.</returns>
    internal Task<Item?> UpdateItem(
        Item item,
        FolderId savedItemsDestinationFolderId,
        ConflictResolutionMode conflictResolution,
        MessageDisposition? messageDisposition,
        SendInvitationsOrCancellationsMode? sendInvitationsOrCancellationsMode,
        CancellationToken token
    )
    {
        return UpdateItem(
            item,
            savedItemsDestinationFolderId,
            conflictResolution,
            messageDisposition,
            sendInvitationsOrCancellationsMode,
            false,
            token
        );
    }

    /// <summary>
    ///     Updates an item.
    /// </summary>
    /// <param name="item">The item to update.</param>
    /// <param name="savedItemsDestinationFolderId">
    ///     The folder in which to save sent messages, meeting invitations or
    ///     cancellations. If null, the message, meeting invitation or cancellation is saved in the Sent Items folder.
    /// </param>
    /// <param name="conflictResolution">The conflict resolution mode.</param>
    /// <param name="messageDisposition">
    ///     Indicates the disposition mode for an item of type EmailMessage. Required if item is
    ///     an EmailMessage instance.
    /// </param>
    /// <param name="sendInvitationsOrCancellationsMode">
    ///     Indicates if and how invitations and/or cancellations should be sent
    ///     for ian tem of type Appointment. Required if item is an Appointment instance.
    /// </param>
    /// <param name="suppressReadReceipts">Whether to suppress read receipts</param>
    /// <param name="token"></param>
    /// <returns>Updated item.</returns>
    internal async Task<Item?> UpdateItem(
        Item item,
        FolderId? savedItemsDestinationFolderId,
        ConflictResolutionMode conflictResolution,
        MessageDisposition? messageDisposition,
        SendInvitationsOrCancellationsMode? sendInvitationsOrCancellationsMode,
        bool suppressReadReceipts,
        CancellationToken token
    )
    {
        var responses = await InternalUpdateItems(
                [item,],
                savedItemsDestinationFolderId,
                conflictResolution,
                messageDisposition,
                sendInvitationsOrCancellationsMode,
                ServiceErrorHandling.ThrowOnError,
                suppressReadReceipts,
                token
            )
            .ConfigureAwait(false);

        return responses[0].ReturnedItem;
    }

    /// <summary>
    ///     Sends an item.
    /// </summary>
    /// <param name="item">The item.</param>
    /// <param name="savedCopyDestinationFolderId">The saved copy destination folder id.</param>
    /// <param name="token"></param>
    internal System.Threading.Tasks.Task SendItem(
        Item item,
        FolderId savedCopyDestinationFolderId,
        CancellationToken token
    )
    {
        var request = new SendItemRequest(this, ServiceErrorHandling.ThrowOnError)
        {
            Items = [item,],
            SavedCopyDestinationFolderId = savedCopyDestinationFolderId,
        };

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Copies multiple items in a single call to EWS.
    /// </summary>
    /// <param name="itemIds">The Ids of the items to copy.</param>
    /// <param name="destinationFolderId">The Id of the folder to copy the items to.</param>
    /// <param name="returnNewItemIds">Flag indicating whether service should return new ItemIds or not.</param>
    /// <param name="errorHandling">What type of error handling should be performed.</param>
    /// <param name="token"></param>
    /// <returns>A ServiceResponseCollection providing copy results for each of the specified item Ids.</returns>
    private Task<ServiceResponseCollection<MoveCopyItemResponse>> InternalCopyItems(
        IEnumerable<ItemId> itemIds,
        FolderId destinationFolderId,
        bool? returnNewItemIds,
        ServiceErrorHandling errorHandling,
        CancellationToken token
    )
    {
        var request = new CopyItemRequest(this, errorHandling);
        request.ItemIds.AddRange(itemIds);
        request.DestinationFolderId = destinationFolderId;
        request.ReturnNewItemIds = returnNewItemIds;

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Copies multiple items in a single call to EWS.
    /// </summary>
    /// <param name="itemIds">The Ids of the items to copy.</param>
    /// <param name="destinationFolderId">The Id of the folder to copy the items to.</param>
    /// <param name="token"></param>
    /// <returns>A ServiceResponseCollection providing copy results for each of the specified item Ids.</returns>
    public Task<ServiceResponseCollection<MoveCopyItemResponse>> CopyItems(
        IEnumerable<ItemId> itemIds,
        FolderId destinationFolderId,
        CancellationToken token = default
    )
    {
        return InternalCopyItems(itemIds, destinationFolderId, null, ServiceErrorHandling.ReturnErrors, token);
    }

    /// <summary>
    ///     Copies multiple items in a single call to EWS.
    /// </summary>
    /// <param name="itemIds">The Ids of the items to copy.</param>
    /// <param name="destinationFolderId">The Id of the folder to copy the items to.</param>
    /// <param name="returnNewItemIds">Flag indicating whether service should return new ItemIds or not.</param>
    /// <param name="token"></param>
    /// <returns>A ServiceResponseCollection providing copy results for each of the specified item Ids.</returns>
    public Task<ServiceResponseCollection<MoveCopyItemResponse>> CopyItems(
        IEnumerable<ItemId> itemIds,
        FolderId destinationFolderId,
        bool returnNewItemIds,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2010_SP1, "CopyItems");

        return InternalCopyItems(
            itemIds,
            destinationFolderId,
            returnNewItemIds,
            ServiceErrorHandling.ReturnErrors,
            token
        );
    }

    /// <summary>
    ///     Copies an item. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="itemId">The Id of the item to copy.</param>
    /// <param name="destinationFolderId">The Id of the folder to copy the item to.</param>
    /// <param name="token"></param>
    /// <returns>The copy of the item.</returns>
    internal async Task<Item> CopyItem(ItemId itemId, FolderId destinationFolderId, CancellationToken token)
    {
        var result = await InternalCopyItems(
                [itemId,],
                destinationFolderId,
                null,
                ServiceErrorHandling.ThrowOnError,
                token
            )
            .ConfigureAwait(false);

        return result[0].Item;
    }

    /// <summary>
    ///     Moves multiple items in a single call to EWS.
    /// </summary>
    /// <param name="itemIds">The Ids of the items to move.</param>
    /// <param name="destinationFolderId">The Id of the folder to move the items to.</param>
    /// <param name="returnNewItemIds">Flag indicating whether service should return new ItemIds or not.</param>
    /// <param name="errorHandling">What type of error handling should be performed.</param>
    /// <param name="token"></param>
    /// <returns>A ServiceResponseCollection providing copy results for each of the specified item Ids.</returns>
    private Task<ServiceResponseCollection<MoveCopyItemResponse>> InternalMoveItems(
        IEnumerable<ItemId> itemIds,
        FolderId destinationFolderId,
        bool? returnNewItemIds,
        ServiceErrorHandling errorHandling,
        CancellationToken token
    )
    {
        var request = new MoveItemRequest(this, errorHandling)
        {
            DestinationFolderId = destinationFolderId,
            ReturnNewItemIds = returnNewItemIds,
        };

        request.ItemIds.AddRange(itemIds);

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Moves multiple items in a single call to EWS.
    /// </summary>
    /// <param name="itemIds">The Ids of the items to move.</param>
    /// <param name="destinationFolderId">The Id of the folder to move the items to.</param>
    /// <param name="token"></param>
    /// <returns>A ServiceResponseCollection providing copy results for each of the specified item Ids.</returns>
    public Task<ServiceResponseCollection<MoveCopyItemResponse>> MoveItems(
        IEnumerable<ItemId> itemIds,
        FolderId destinationFolderId,
        CancellationToken token = default
    )
    {
        return InternalMoveItems(itemIds, destinationFolderId, null, ServiceErrorHandling.ReturnErrors, token);
    }

    /// <summary>
    ///     Moves multiple items in a single call to EWS.
    /// </summary>
    /// <param name="itemIds">The Ids of the items to move.</param>
    /// <param name="destinationFolderId">The Id of the folder to move the items to.</param>
    /// <param name="returnNewItemIds">Flag indicating whether service should return new ItemIds or not.</param>
    /// <param name="token"></param>
    /// <returns>A ServiceResponseCollection providing copy results for each of the specified item Ids.</returns>
    public Task<ServiceResponseCollection<MoveCopyItemResponse>> MoveItems(
        IEnumerable<ItemId> itemIds,
        FolderId destinationFolderId,
        bool returnNewItemIds,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2010_SP1, "MoveItems");

        return InternalMoveItems(
            itemIds,
            destinationFolderId,
            returnNewItemIds,
            ServiceErrorHandling.ReturnErrors,
            token
        );
    }

    /// <summary>
    ///     Move an item.
    /// </summary>
    /// <param name="itemId">The Id of the item to move.</param>
    /// <param name="destinationFolderId">The Id of the folder to move the item to.</param>
    /// <param name="token"></param>
    /// <returns>The moved item.</returns>
    internal async Task<Item> MoveItem(ItemId itemId, FolderId destinationFolderId, CancellationToken token)
    {
        var result = await InternalMoveItems(
                [itemId,],
                destinationFolderId,
                null,
                ServiceErrorHandling.ThrowOnError,
                token
            )
            .ConfigureAwait(false);

        return result[0].Item;
    }

    /// <summary>
    ///     Archives multiple items in a single call to EWS.
    /// </summary>
    /// <param name="itemIds">The Ids of the items to move.</param>
    /// <param name="sourceFolderId">The Id of the folder in primary corresponding to which items are being archived to.</param>
    /// <param name="token"></param>
    /// <returns>A ServiceResponseCollection providing copy results for each of the specified item Ids.</returns>
    public Task<ServiceResponseCollection<ArchiveItemResponse>> ArchiveItems(
        IEnumerable<ItemId> itemIds,
        FolderId sourceFolderId,
        CancellationToken token = default
    )
    {
        var request = new ArchiveItemRequest(this, ServiceErrorHandling.ReturnErrors)
        {
            SourceFolderId = sourceFolderId,
        };

        request.Ids.AddRange(itemIds);

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Finds items.
    /// </summary>
    /// <typeparam name="TItem">The type of the item.</typeparam>
    /// <param name="parentFolderIds">The parent folder ids.</param>
    /// <param name="searchFilter">
    ///     The search filter. Available search filter classes
    ///     include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and
    ///     SearchFilter.SearchFilterCollection
    /// </param>
    /// <param name="queryString">query string to be used for indexed search.</param>
    /// <param name="view">The view controlling the number of items returned.</param>
    /// <param name="groupBy">The group by.</param>
    /// <param name="errorHandlingMode">Indicates the type of error handling should be done.</param>
    /// <param name="token"></param>
    /// <returns>Service response collection.</returns>
    internal Task<ServiceResponseCollection<FindItemResponse<TItem>>> FindItems<TItem>(
        IEnumerable<FolderId> parentFolderIds,
        SearchFilter? searchFilter,
        string? queryString,
        ViewBase view,
        Grouping? groupBy,
        ServiceErrorHandling errorHandlingMode,
        CancellationToken token
    )
        where TItem : Item
    {
        EwsUtilities.ValidateParamCollection(parentFolderIds);
        EwsUtilities.ValidateParam(view);
        EwsUtilities.ValidateParamAllowNull(groupBy);
        EwsUtilities.ValidateParamAllowNull(queryString);
        EwsUtilities.ValidateParamAllowNull(searchFilter);

        var request = new FindItemRequest<TItem>(this, errorHandlingMode)
        {
            SearchFilter = searchFilter,
            QueryString = queryString,
            View = view,
            GroupBy = groupBy,
        };

        request.ParentFolderIds.AddRange(parentFolderIds);

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Obtains a list of items by searching the contents of a specific folder. Calling this method results in a call to
    ///     EWS.
    /// </summary>
    /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
    /// <param name="queryString">the search string to be used for indexed search, if any.</param>
    /// <param name="view">The view controlling the number of items returned.</param>
    /// <param name="token"></param>
    /// <returns>An object representing the results of the search operation.</returns>
    public async Task<FindItemsResults<Item>> FindItems(
        FolderId parentFolderId,
        string queryString,
        ViewBase view,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParamAllowNull(queryString);

        var responses = await FindItems<Item>(
                [parentFolderId,],
                null,
                queryString,
                view,
                null,
                ServiceErrorHandling.ThrowOnError,
                token
            )
            .ConfigureAwait(false);

        return responses[0].Results;
    }

    /// <summary>
    ///     Obtains a list of items by searching the contents of a specific folder.
    ///     Along with conversations, a list of highlight terms are returned.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
    /// <param name="queryString">the search string to be used for indexed search, if any.</param>
    /// <param name="returnHighlightTerms">Flag indicating if highlight terms should be returned in the response</param>
    /// <param name="view">The view controlling the number of items returned.</param>
    /// <param name="token"></param>
    /// <returns>An object representing the results of the search operation.</returns>
    public async Task<FindItemsResults<Item>> FindItems(
        FolderId parentFolderId,
        string queryString,
        bool returnHighlightTerms,
        ViewBase view,
        CancellationToken token = default
    )
    {
        FolderId[] parentFolderIds = [parentFolderId,];

        EwsUtilities.ValidateParamCollection(parentFolderIds);
        EwsUtilities.ValidateParam(view);
        EwsUtilities.ValidateParamAllowNull(queryString);
        EwsUtilities.ValidateParamAllowNull(returnHighlightTerms);
        EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013, "FindItems");

        var request = new FindItemRequest<Item>(this, ServiceErrorHandling.ThrowOnError);

        request.ParentFolderIds.AddRange(parentFolderIds);
        request.QueryString = queryString;
        request.ReturnHighlightTerms = returnHighlightTerms;
        request.View = view;

        var responses = await request.ExecuteAsync(token).ConfigureAwait(false);
        return responses[0].Results;
    }

    /// <summary>
    ///     Obtains a list of items by searching the contents of a specific folder.
    ///     Along with conversations, a list of highlight terms are returned.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
    /// <param name="queryString">the search string to be used for indexed search, if any.</param>
    /// <param name="returnHighlightTerms">Flag indicating if highlight terms should be returned in the response</param>
    /// <param name="view">The view controlling the number of items returned.</param>
    /// <param name="groupBy">The group by clause.</param>
    /// <param name="token"></param>
    /// <returns>An object representing the results of the search operation.</returns>
    public async Task<GroupedFindItemsResults<Item>> FindItems(
        FolderId parentFolderId,
        string queryString,
        bool returnHighlightTerms,
        ViewBase view,
        Grouping groupBy,
        CancellationToken token = default
    )
    {
        FolderId[] parentFolderIds = [parentFolderId,];

        EwsUtilities.ValidateParamCollection(parentFolderIds);
        EwsUtilities.ValidateParam(view);
        EwsUtilities.ValidateParam(groupBy);
        EwsUtilities.ValidateParamAllowNull(queryString);
        EwsUtilities.ValidateParamAllowNull(returnHighlightTerms);
        EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013, "FindItems");

        var request = new FindItemRequest<Item>(this, ServiceErrorHandling.ThrowOnError)
        {
            QueryString = queryString,
            ReturnHighlightTerms = returnHighlightTerms,
            View = view,
            GroupBy = groupBy,
        };

        request.ParentFolderIds.AddRange(parentFolderIds);

        var responses = await request.ExecuteAsync(token).ConfigureAwait(false);
        return responses[0].GroupedFindResults;
    }

    /// <summary>
    ///     Obtains a list of items by searching the contents of a specific folder. Calling this method results in a call to
    ///     EWS.
    /// </summary>
    /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
    /// <param name="searchFilter">
    ///     The search filter. Available search filter classes
    ///     include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and
    ///     SearchFilter.SearchFilterCollection
    /// </param>
    /// <param name="view">The view controlling the number of items returned.</param>
    /// <param name="token"></param>
    /// <returns>An object representing the results of the search operation.</returns>
    public async Task<FindItemsResults<Item>> FindItems(
        FolderId parentFolderId,
        SearchFilter? searchFilter,
        ViewBase view,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParamAllowNull(searchFilter);

        var responses = await FindItems<Item>(
                [parentFolderId,],
                searchFilter,
                null,
                view,
                null,
                ServiceErrorHandling.ThrowOnError,
                token
            )
            .ConfigureAwait(false);

        return responses[0].Results;
    }

    /// <summary>
    ///     Obtains a list of items by searching the contents of a specific folder. Calling this method results in a call to
    ///     EWS.
    /// </summary>
    /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
    /// <param name="view">The view controlling the number of items returned.</param>
    /// <param name="token"></param>
    /// <returns>An object representing the results of the search operation.</returns>
    public async Task<FindItemsResults<Item>> FindItems(
        FolderId parentFolderId,
        ViewBase view,
        CancellationToken token = default
    )
    {
        var responses = await FindItems<Item>(
                new[]
                {
                    parentFolderId,
                },
                null,
                null,
                view,
                null,
                ServiceErrorHandling.ThrowOnError,
                token
            )
            .ConfigureAwait(false);

        return responses[0].Results;
    }

    /// <summary>
    ///     Obtains a list of items by searching the contents of a specific folder. Calling this method results in a call to
    ///     EWS.
    /// </summary>
    /// <param name="parentFolderName">The name of the folder in which to search for items.</param>
    /// <param name="queryString">query string to be used for indexed search</param>
    /// <param name="view">The view controlling the number of items returned.</param>
    /// <returns>An object representing the results of the search operation.</returns>
    public Task<FindItemsResults<Item>> FindItems(
        WellKnownFolderName parentFolderName,
        string queryString,
        ViewBase view
    )
    {
        return FindItems(new FolderId(parentFolderName), queryString, view);
    }

    /// <summary>
    ///     Obtains a list of items by searching the contents of a specific folder. Calling this method results in a call to
    ///     EWS.
    /// </summary>
    /// <param name="parentFolderName">The name of the folder in which to search for items.</param>
    /// <param name="searchFilter">
    ///     The search filter. Available search filter classes
    ///     include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and
    ///     SearchFilter.SearchFilterCollection
    /// </param>
    /// <param name="view">The view controlling the number of items returned.</param>
    /// <returns>An object representing the results of the search operation.</returns>
    public Task<FindItemsResults<Item>> FindItems(
        WellKnownFolderName parentFolderName,
        SearchFilter searchFilter,
        ViewBase view
    )
    {
        return FindItems(new FolderId(parentFolderName), searchFilter, view);
    }

    /// <summary>
    ///     Obtains a list of items by searching the contents of a specific folder. Calling this method results in a call to
    ///     EWS.
    /// </summary>
    /// <param name="parentFolderName">The name of the folder in which to search for items.</param>
    /// <param name="view">The view controlling the number of items returned.</param>
    /// <param name="token"></param>
    /// <returns>An object representing the results of the search operation.</returns>
    public Task<FindItemsResults<Item>> FindItems(
        WellKnownFolderName parentFolderName,
        ViewBase view,
        CancellationToken token = default
    )
    {
        return FindItems(new FolderId(parentFolderName), (SearchFilter?)null, view, token);
    }

    /// <summary>
    ///     Obtains a grouped list of items by searching the contents of a specific folder. Calling this method results in a
    ///     call to EWS.
    /// </summary>
    /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
    /// <param name="queryString">query string to be used for indexed search</param>
    /// <param name="view">The view controlling the number of items returned.</param>
    /// <param name="groupBy">The group by clause.</param>
    /// <param name="token"></param>
    /// <returns>A list of items containing the contents of the specified folder.</returns>
    public async Task<GroupedFindItemsResults<Item>> FindItems(
        FolderId parentFolderId,
        string queryString,
        ViewBase view,
        Grouping groupBy,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(groupBy);
        EwsUtilities.ValidateParamAllowNull(queryString);

        var responses = await FindItems<Item>(
                new[]
                {
                    parentFolderId,
                },
                null,
                queryString,
                view,
                groupBy,
                ServiceErrorHandling.ThrowOnError,
                token
            )
            .ConfigureAwait(false);

        return responses[0].GroupedFindResults;
    }

    /// <summary>
    ///     Obtains a grouped list of items by searching the contents of a specific folder. Calling this method results in a
    ///     call to EWS.
    /// </summary>
    /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
    /// <param name="searchFilter">
    ///     The search filter. Available search filter classes
    ///     include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and
    ///     SearchFilter.SearchFilterCollection
    /// </param>
    /// <param name="view">The view controlling the number of items returned.</param>
    /// <param name="groupBy">The group by clause.</param>
    /// <param name="token"></param>
    /// <returns>A list of items containing the contents of the specified folder.</returns>
    public async Task<GroupedFindItemsResults<Item>> FindItems(
        FolderId parentFolderId,
        SearchFilter searchFilter,
        ViewBase view,
        Grouping groupBy,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(groupBy);
        EwsUtilities.ValidateParamAllowNull(searchFilter);

        var responses = await FindItems<Item>(
                new[]
                {
                    parentFolderId,
                },
                searchFilter,
                null,
                view,
                groupBy,
                ServiceErrorHandling.ThrowOnError,
                token
            )
            .ConfigureAwait(false);

        return responses[0].GroupedFindResults;
    }

    /// <summary>
    ///     Obtains a grouped list of items by searching the contents of a specific folder. Calling this method results in a
    ///     call to EWS.
    /// </summary>
    /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
    /// <param name="view">The view controlling the number of items returned.</param>
    /// <param name="groupBy">The group by clause.</param>
    /// <param name="token"></param>
    /// <returns>A list of items containing the contents of the specified folder.</returns>
    public async Task<GroupedFindItemsResults<Item>> FindItems(
        FolderId parentFolderId,
        ViewBase view,
        Grouping groupBy,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(groupBy);

        var responses = await FindItems<Item>(
                new[]
                {
                    parentFolderId,
                },
                null,
                null,
                view,
                groupBy,
                ServiceErrorHandling.ThrowOnError,
                token
            )
            .ConfigureAwait(false);

        return responses[0].GroupedFindResults;
    }

    /// <summary>
    ///     Obtains a grouped list of items by searching the contents of a specific folder. Calling this method results in a
    ///     call to EWS.
    /// </summary>
    /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
    /// <param name="searchFilter">
    ///     The search filter. Available search filter classes
    ///     include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and
    ///     SearchFilter.SearchFilterCollection
    /// </param>
    /// <param name="view">The view controlling the number of items returned.</param>
    /// <param name="groupBy">The group by clause.</param>
    /// <param name="token"></param>
    /// <typeparam name="TItem">Type of item.</typeparam>
    /// <returns>A list of items containing the contents of the specified folder.</returns>
    internal Task<ServiceResponseCollection<FindItemResponse<TItem>>> FindItems<TItem>(
        FolderId parentFolderId,
        SearchFilter searchFilter,
        ViewBase view,
        Grouping groupBy,
        CancellationToken token
    )
        where TItem : Item
    {
        return FindItems<TItem>(
            new[]
            {
                parentFolderId,
            },
            searchFilter,
            null,
            view,
            groupBy,
            ServiceErrorHandling.ThrowOnError,
            token
        );
    }

    /// <summary>
    ///     Obtains a grouped list of items by searching the contents of a specific folder. Calling this method results in a
    ///     call to EWS.
    /// </summary>
    /// <param name="parentFolderName">The name of the folder in which to search for items.</param>
    /// <param name="queryString">query string to be used for indexed search</param>
    /// <param name="view">The view controlling the number of items returned.</param>
    /// <param name="groupBy">The group by clause.</param>
    /// <param name="token"></param>
    /// <returns>A collection of grouped items representing the contents of the specified.</returns>
    public Task<GroupedFindItemsResults<Item>> FindItems(
        WellKnownFolderName parentFolderName,
        string queryString,
        ViewBase view,
        Grouping groupBy,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(groupBy);

        return FindItems(new FolderId(parentFolderName), queryString, view, groupBy, token);
    }

    /// <summary>
    ///     Obtains a grouped list of items by searching the contents of a specific folder. Calling this method results in a
    ///     call to EWS.
    /// </summary>
    /// <param name="parentFolderName">The name of the folder in which to search for items.</param>
    /// <param name="searchFilter">
    ///     The search filter. Available search filter classes
    ///     include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and
    ///     SearchFilter.SearchFilterCollection
    /// </param>
    /// <param name="view">The view controlling the number of items returned.</param>
    /// <param name="groupBy">The group by clause.</param>
    /// <param name="token"></param>
    /// <returns>A collection of grouped items representing the contents of the specified.</returns>
    public Task<GroupedFindItemsResults<Item>> FindItems(
        WellKnownFolderName parentFolderName,
        SearchFilter searchFilter,
        ViewBase view,
        Grouping groupBy,
        CancellationToken token = default
    )
    {
        return FindItems(new FolderId(parentFolderName), searchFilter, view, groupBy, token);
    }

    /// <summary>
    ///     Obtains a list of appointments by searching the contents of a specific folder. Calling this method results in a
    ///     call to EWS.
    /// </summary>
    /// <param name="parentFolderId">The id of the calendar folder in which to search for items.</param>
    /// <param name="calendarView">The calendar view controlling the number of appointments returned.</param>
    /// <param name="token"></param>
    /// <returns>A collection of appointments representing the contents of the specified folder.</returns>
    public async Task<FindItemsResults<Appointment>> FindAppointments(
        FolderId parentFolderId,
        CalendarView calendarView,
        CancellationToken token = default
    )
    {
        var response = await FindItems<Appointment>(
                new[]
                {
                    parentFolderId,
                },
                null,
                null,
                calendarView,
                null,
                ServiceErrorHandling.ThrowOnError,
                token
            )
            .ConfigureAwait(false);

        return response[0].Results;
    }

    /// <summary>
    ///     Obtains a list of appointments by searching the contents of a specific folder. Calling this method results in a
    ///     call to EWS.
    /// </summary>
    /// <param name="parentFolderName">The name of the calendar folder in which to search for items.</param>
    /// <param name="calendarView">The calendar view controlling the number of appointments returned.</param>
    /// <param name="token"></param>
    /// <returns>A collection of appointments representing the contents of the specified folder.</returns>
    public Task<FindItemsResults<Appointment>> FindAppointments(
        WellKnownFolderName parentFolderName,
        CalendarView calendarView,
        CancellationToken token = default
    )
    {
        return FindAppointments(new FolderId(parentFolderName), calendarView, token);
    }

    /// <summary>
    ///     Loads the properties of multiple items in a single call to EWS.
    /// </summary>
    /// <param name="items">The items to load the properties of.</param>
    /// <param name="propertySet">The set of properties to load.</param>
    /// <param name="token"></param>
    /// <returns>A ServiceResponseCollection providing results for each of the specified items.</returns>
    public Task<ServiceResponseCollection<ServiceResponse>> LoadPropertiesForItems(
        IEnumerable<Item> items,
        PropertySet propertySet,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParamCollection(items);
        EwsUtilities.ValidateParam(propertySet);

        return InternalLoadPropertiesForItems(items, propertySet, ServiceErrorHandling.ReturnErrors, token);
    }

    /// <summary>
    ///     Loads the properties of multiple items in a single call to EWS.
    /// </summary>
    /// <param name="items">The items to load the properties of.</param>
    /// <param name="propertySet">The set of properties to load.</param>
    /// <param name="errorHandling">Indicates the type of error handling should be done.</param>
    /// <param name="token"></param>
    /// <returns>A ServiceResponseCollection providing results for each of the specified items.</returns>
    internal Task<ServiceResponseCollection<ServiceResponse>> InternalLoadPropertiesForItems(
        IEnumerable<Item> items,
        PropertySet propertySet,
        ServiceErrorHandling errorHandling,
        CancellationToken token
    )
    {
        var request = new GetItemRequestForLoad(this, errorHandling)
        {
            PropertySet = propertySet,
        };

        request.ItemIds.AddRange(items);

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Binds to multiple items in a single call to EWS.
    /// </summary>
    /// <param name="itemIds">The Ids of the items to bind to.</param>
    /// <param name="propertySet">The set of properties to load.</param>
    /// <param name="anchorMailbox">The SmtpAddress of mailbox that hosts all items we need to bind to</param>
    /// <param name="errorHandling">Type of error handling to perform.</param>
    /// <param name="token"></param>
    /// <returns>A ServiceResponseCollection providing results for each of the specified item Ids.</returns>
    private Task<ServiceResponseCollection<GetItemResponse>> InternalBindToItems(
        IEnumerable<ItemId> itemIds,
        PropertySet propertySet,
        string? anchorMailbox,
        ServiceErrorHandling errorHandling,
        CancellationToken token
    )
    {
        var request = new GetItemRequest(this, errorHandling)
        {
            PropertySet = propertySet,
            AnchorMailbox = anchorMailbox,
        };

        request.ItemIds.AddRange(itemIds);

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Binds to multiple items in a single call to EWS.
    /// </summary>
    /// <param name="itemIds">The Ids of the items to bind to.</param>
    /// <param name="propertySet">The set of properties to load.</param>
    /// <param name="token"></param>
    /// <returns>A ServiceResponseCollection providing results for each of the specified item Ids.</returns>
    public Task<ServiceResponseCollection<GetItemResponse>> BindToItems(
        IEnumerable<ItemId> itemIds,
        PropertySet propertySet,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParamCollection(itemIds);
        EwsUtilities.ValidateParam(propertySet);

        return InternalBindToItems(itemIds, propertySet, null, ServiceErrorHandling.ReturnErrors, token);
    }

    /// <summary>
    ///     Binds to multiple items in a single call to EWS.
    /// </summary>
    /// <param name="itemIds">The Ids of the items to bind to.</param>
    /// <param name="propertySet">The set of properties to load.</param>
    /// <param name="anchorMailbox">The SmtpAddress of mailbox that hosts all items we need to bind to</param>
    /// <param name="token"></param>
    /// <returns>A ServiceResponseCollection providing results for each of the specified item Ids.</returns>
    /// <remarks>
    ///     This API designed to be used primarily in groups scenarios where we want to set the
    ///     anchor mailbox header so that request is routed directly to the group mailbox backend server.
    /// </remarks>
    public Task<ServiceResponseCollection<GetItemResponse>> BindToGroupItems(
        IEnumerable<ItemId> itemIds,
        PropertySet propertySet,
        string anchorMailbox,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParamCollection(itemIds);
        EwsUtilities.ValidateParam(propertySet);
        EwsUtilities.ValidateParam(anchorMailbox);

        return InternalBindToItems(itemIds, propertySet, anchorMailbox, ServiceErrorHandling.ReturnErrors, token);
    }

    /// <summary>
    ///     Binds to item.
    /// </summary>
    /// <param name="itemId">The item id.</param>
    /// <param name="propertySet">The property set.</param>
    /// <param name="token"></param>
    /// <returns>Item.</returns>
    internal async Task<Item?> BindToItem(ItemId itemId, PropertySet propertySet, CancellationToken token)
    {
        EwsUtilities.ValidateParam(itemId);
        EwsUtilities.ValidateParam(propertySet);

        var responses = await InternalBindToItems(
                new[]
                {
                    itemId,
                },
                propertySet,
                null,
                ServiceErrorHandling.ThrowOnError,
                token
            )
            .ConfigureAwait(false);

        return responses[0].Item;
    }

    /// <summary>
    ///     Binds to item.
    /// </summary>
    /// <typeparam name="TItem">The type of the item.</typeparam>
    /// <param name="itemId">The item id.</param>
    /// <param name="propertySet">The property set.</param>
    /// <param name="token"></param>
    /// <returns>Item</returns>
    internal async Task<TItem> BindToItem<TItem>(ItemId itemId, PropertySet propertySet, CancellationToken token)
        where TItem : Item
    {
        var result = await BindToItem(itemId, propertySet, token).ConfigureAwait(false);

        if (result is TItem item)
        {
            return item;
        }

        throw new ServiceLocalException(
            string.Format(Strings.ItemTypeNotCompatible, result.GetType().Name, typeof(TItem).Name)
        );
    }

    /// <summary>
    ///     Deletes multiple items in a single call to EWS.
    /// </summary>
    /// <param name="itemIds">The Ids of the items to delete.</param>
    /// <param name="deleteMode">The deletion mode.</param>
    /// <param name="sendCancellationsMode">
    ///     Indicates whether cancellation messages should be sent. Required if any of the item
    ///     Ids represents an Appointment.
    /// </param>
    /// <param name="affectedTaskOccurrences">
    ///     Indicates which instance of a recurring task should be deleted. Required if any
    ///     of the item Ids represents a Task.
    /// </param>
    /// <param name="errorHandling">Type of error handling to perform.</param>
    /// <param name="suppressReadReceipts">Whether to suppress read receipts</param>
    /// <param name="token"></param>
    /// <returns>A ServiceResponseCollection providing deletion results for each of the specified item Ids.</returns>
    private Task<ServiceResponseCollection<ServiceResponse>> InternalDeleteItems(
        IEnumerable<ItemId> itemIds,
        DeleteMode deleteMode,
        SendCancellationsMode? sendCancellationsMode,
        AffectedTaskOccurrence? affectedTaskOccurrences,
        ServiceErrorHandling errorHandling,
        bool suppressReadReceipts,
        CancellationToken token
    )
    {
        var request = new DeleteItemRequest(this, errorHandling)
        {
            DeleteMode = deleteMode,
            SendCancellationsMode = sendCancellationsMode,
            AffectedTaskOccurrences = affectedTaskOccurrences,
            SuppressReadReceipts = suppressReadReceipts,
        };

        request.ItemIds.AddRange(itemIds);

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Deletes multiple items in a single call to EWS.
    /// </summary>
    /// <param name="itemIds">The Ids of the items to delete.</param>
    /// <param name="deleteMode">The deletion mode.</param>
    /// <param name="sendCancellationsMode">
    ///     Indicates whether cancellation messages should be sent. Required if any of the item
    ///     Ids represents an Appointment.
    /// </param>
    /// <param name="affectedTaskOccurrences">
    ///     Indicates which instance of a recurring task should be deleted. Required if any
    ///     of the item Ids represents a Task.
    /// </param>
    /// <returns>A ServiceResponseCollection providing deletion results for each of the specified item Ids.</returns>
    public Task<ServiceResponseCollection<ServiceResponse>> DeleteItems(
        IEnumerable<ItemId> itemIds,
        DeleteMode deleteMode,
        SendCancellationsMode? sendCancellationsMode,
        AffectedTaskOccurrence? affectedTaskOccurrences
    )
    {
        return DeleteItems(itemIds, deleteMode, sendCancellationsMode, affectedTaskOccurrences, false);
    }

    /// <summary>
    ///     Deletes multiple items in a single call to EWS.
    /// </summary>
    /// <param name="itemIds">The Ids of the items to delete.</param>
    /// <param name="deleteMode">The deletion mode.</param>
    /// <param name="sendCancellationsMode">
    ///     Indicates whether cancellation messages should be sent. Required if any of the item
    ///     Ids represents an Appointment.
    /// </param>
    /// <param name="affectedTaskOccurrences">
    ///     Indicates which instance of a recurring task should be deleted. Required if any
    ///     of the item Ids represents a Task.
    /// </param>
    /// <returns>A ServiceResponseCollection providing deletion results for each of the specified item Ids.</returns>
    /// <param name="suppressReadReceipt">Whether to suppress read receipts</param>
    /// <param name="token"></param>
    public Task<ServiceResponseCollection<ServiceResponse>> DeleteItems(
        IEnumerable<ItemId> itemIds,
        DeleteMode deleteMode,
        SendCancellationsMode? sendCancellationsMode,
        AffectedTaskOccurrence? affectedTaskOccurrences,
        bool suppressReadReceipt,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParamCollection(itemIds);

        return InternalDeleteItems(
            itemIds,
            deleteMode,
            sendCancellationsMode,
            affectedTaskOccurrences,
            ServiceErrorHandling.ReturnErrors,
            suppressReadReceipt,
            token
        );
    }

    /// <summary>
    ///     Deletes an item. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="itemId">The Id of the item to delete.</param>
    /// <param name="deleteMode">The deletion mode.</param>
    /// <param name="sendCancellationsMode">
    ///     Indicates whether cancellation messages should be sent. Required if the item Id
    ///     represents an Appointment.
    /// </param>
    /// <param name="affectedTaskOccurrences">
    ///     Indicates which instance of a recurring task should be deleted. Required if item
    ///     Id represents a Task.
    /// </param>
    /// <param name="token"></param>
    internal Task<ServiceResponseCollection<ServiceResponse>> DeleteItem(
        ItemId itemId,
        DeleteMode deleteMode,
        SendCancellationsMode? sendCancellationsMode,
        AffectedTaskOccurrence? affectedTaskOccurrences,
        CancellationToken token
    )
    {
        return DeleteItem(itemId, deleteMode, sendCancellationsMode, affectedTaskOccurrences, false, token);
    }

    /// <summary>
    ///     Deletes an item. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="itemId">The Id of the item to delete.</param>
    /// <param name="deleteMode">The deletion mode.</param>
    /// <param name="sendCancellationsMode">
    ///     Indicates whether cancellation messages should be sent. Required if the item Id
    ///     represents an Appointment.
    /// </param>
    /// <param name="affectedTaskOccurrences">
    ///     Indicates which instance of a recurring task should be deleted. Required if item
    ///     Id represents a Task.
    /// </param>
    /// <param name="suppressReadReceipts">Whether to suppress read receipts</param>
    /// <param name="token"></param>
    internal Task<ServiceResponseCollection<ServiceResponse>> DeleteItem(
        ItemId itemId,
        DeleteMode deleteMode,
        SendCancellationsMode? sendCancellationsMode,
        AffectedTaskOccurrence? affectedTaskOccurrences,
        bool suppressReadReceipts,
        CancellationToken token
    )
    {
        EwsUtilities.ValidateParam(itemId);

        return InternalDeleteItems(
            new[]
            {
                itemId,
            },
            deleteMode,
            sendCancellationsMode,
            affectedTaskOccurrences,
            ServiceErrorHandling.ThrowOnError,
            suppressReadReceipts,
            token
        );
    }

    /// <summary>
    ///     Mark items as junk.
    /// </summary>
    /// <param name="itemIds">ItemIds for the items to mark</param>
    /// <param name="isJunk">
    ///     Whether the items are junk.  If true, senders are add to blocked sender list. If false, senders
    ///     are removed.
    /// </param>
    /// <param name="moveItem">
    ///     Whether to move the item.  Items are moved to junk folder if isJunk is true, inbox if isJunk is
    ///     false.
    /// </param>
    /// <param name="token"></param>
    /// <returns>A ServiceResponseCollection providing itemIds for each of the moved items..</returns>
    public Task<ServiceResponseCollection<MarkAsJunkResponse>> MarkAsJunk(
        IEnumerable<ItemId> itemIds,
        bool isJunk,
        bool moveItem,
        CancellationToken token = default
    )
    {
        var request = new MarkAsJunkRequest(this, ServiceErrorHandling.ReturnErrors)
        {
            IsJunk = isJunk,
            MoveItem = moveItem,
        };
        request.ItemIds.AddRange(itemIds);

        return request.ExecuteAsync(token);
    }

    #endregion


    #region People operations

    /// <summary>
    ///     This method is for search scenarios. Retrieves a set of personas satisfying the specified search conditions.
    /// </summary>
    /// <param name="folderId">Id of the folder being searched</param>
    /// <param name="searchFilter">
    ///     The search filter. Available search filter classes
    ///     include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and
    ///     SearchFilter.SearchFilterCollection
    /// </param>
    /// <param name="view">The view which defines the number of persona being returned</param>
    /// <param name="queryString">The query string for which the search is being performed</param>
    /// <param name="token"></param>
    /// <returns>A collection of personas matching the search conditions</returns>
    public async Task<ICollection<Persona>> FindPeople(
        FolderId folderId,
        SearchFilter searchFilter,
        ViewBase view,
        string queryString,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParamAllowNull(folderId);
        EwsUtilities.ValidateParamAllowNull(searchFilter);
        EwsUtilities.ValidateParam(view);
        EwsUtilities.ValidateParam(queryString);
        EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013_SP1, "FindPeople");

        var request = new FindPeopleRequest(this)
        {
            FolderId = folderId,
            SearchFilter = searchFilter,
            View = view,
            QueryString = queryString,
        };

        var result = await request.Execute(token).ConfigureAwait(false);
        return result.Personas;
    }

    /// <summary>
    ///     This method is for search scenarios. Retrieves a set of personas satisfying the specified search conditions.
    /// </summary>
    /// <param name="folderName">Name of the folder being searched</param>
    /// <param name="searchFilter">
    ///     The search filter. Available search filter classes
    ///     include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and
    ///     SearchFilter.SearchFilterCollection
    /// </param>
    /// <param name="view">The view which defines the number of persona being returned</param>
    /// <param name="queryString">The query string for which the search is being performed</param>
    /// <returns>A collection of personas matching the search conditions</returns>
    public Task<ICollection<Persona>> FindPeople(
        WellKnownFolderName folderName,
        SearchFilter searchFilter,
        ViewBase view,
        string queryString
    )
    {
        return FindPeople(new FolderId(folderName), searchFilter, view, queryString);
    }

    /// <summary>
    ///     This method is for browse scenarios. Retrieves a set of personas satisfying the specified browse conditions.
    ///     Browse scenariosdon't require query string.
    /// </summary>
    /// <param name="folderId">Id of the folder being browsed</param>
    /// <param name="searchFilter">Search filter</param>
    /// <param name="view">The view which defines paging and the number of persona being returned</param>
    /// <param name="token"></param>
    /// <returns>A result object containing resultset for browsing</returns>
    public async Task<FindPeopleResults> FindPeople(
        FolderId folderId,
        SearchFilter searchFilter,
        ViewBase view,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParamAllowNull(folderId);
        EwsUtilities.ValidateParamAllowNull(searchFilter);
        EwsUtilities.ValidateParamAllowNull(view);
        EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013_SP1, "FindPeople");

        var request = new FindPeopleRequest(this)
        {
            FolderId = folderId,
            SearchFilter = searchFilter,
            View = view,
        };

        var response = await request.Execute(token).ConfigureAwait(false);
        return response.Results;
    }

    /// <summary>
    ///     This method is for browse scenarios. Retrieves a set of personas satisfying the specified browse conditions.
    ///     Browse scenarios don't require query string.
    /// </summary>
    /// <param name="folderName">Name of the folder being browsed</param>
    /// <param name="searchFilter">Search filter</param>
    /// <param name="view">The view which defines paging and the number of personas being returned</param>
    /// <returns>A result object containing resultset for browsing</returns>
    public Task<FindPeopleResults> FindPeople(WellKnownFolderName folderName, SearchFilter searchFilter, ViewBase view)
    {
        return FindPeople(new FolderId(folderName), searchFilter, view);
    }

    /// <summary>
    ///     Retrieves all people who are relevant to the user
    /// </summary>
    /// <param name="view">The view which defines the number of personas being returned</param>
    /// <returns>A collection of personas matching the query string</returns>
    public Task<IPeopleQueryResults> BrowsePeople(ViewBase view)
    {
        return BrowsePeople(view, null);
    }

    /// <summary>
    ///     Retrieves all people who are relevant to the user
    /// </summary>
    /// <param name="view">The view which defines the number of personas being returned</param>
    /// <param name="context">The context for this query. See PeopleQueryContextKeys for keys</param>
    /// <param name="token"></param>
    /// <returns>A collection of personas matching the query string</returns>
    public Task<IPeopleQueryResults> BrowsePeople(
        ViewBase view,
        Dictionary<string, string>? context,
        CancellationToken token = default
    )
    {
        return PerformPeopleQuery(view, string.Empty, context, null, token);
    }

    /// <summary>
    ///     Searches for people who are relevant to the user, automatically determining
    ///     the best sources to use.
    /// </summary>
    /// <param name="view">The view which defines the number of personas being returned</param>
    /// <param name="queryString">The query string for which the search is being performed</param>
    /// <returns>A collection of personas matching the query string</returns>
    public Task<IPeopleQueryResults> SearchPeople(ViewBase view, string queryString)
    {
        return SearchPeople(view, queryString, null, null);
    }

    /// <summary>
    ///     Searches for people who are relevant to the user
    /// </summary>
    /// <param name="view">The view which defines the number of personas being returned</param>
    /// <param name="queryString">The query string for which the search is being performed</param>
    /// <param name="context">The context for this query. See PeopleQueryContextKeys for keys</param>
    /// <param name="queryMode">The scope of the query.</param>
    /// <param name="token"></param>
    /// <returns>A collection of personas matching the query string</returns>
    public Task<IPeopleQueryResults> SearchPeople(
        ViewBase view,
        string queryString,
        Dictionary<string, string>? context,
        PeopleQueryMode? queryMode,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(queryString);

        return PerformPeopleQuery(view, queryString, context, queryMode, token);
    }

    /// <summary>
    ///     Performs a People Query FindPeople call
    /// </summary>
    /// <param name="view">The view which defines the number of personas being returned</param>
    /// <param name="queryString">The query string for which the search is being performed</param>
    /// <param name="context">The context for this query</param>
    /// <param name="queryMode">The scope of the query.</param>
    /// <param name="token"></param>
    /// <returns></returns>
    private async Task<IPeopleQueryResults> PerformPeopleQuery(
        ViewBase view,
        string queryString,
        Dictionary<string, string>? context,
        PeopleQueryMode? queryMode,
        CancellationToken token
    )
    {
        EwsUtilities.ValidateParam(view);
        EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2015, "FindPeople");

        if (context == null)
        {
            context = new Dictionary<string, string>();
        }

        if (queryMode == null)
        {
            queryMode = PeopleQueryMode.Auto;
        }

        var request = new FindPeopleRequest(this)
        {
            View = view,
            QueryString = queryString,
            SearchPeopleSuggestionIndex = true,
            Context = context,
            QueryMode = queryMode,
        };

        var response = await request.Execute(token).ConfigureAwait(false);

        var results = new PeopleQueryResults
        {
            Personas = response.Personas.ToList(),
            TransactionId = response.TransactionId,
        };

        return results;
    }

    /// <summary>
    ///     Get a user's photo.
    /// </summary>
    /// <param name="emailAddress">The user's email address</param>
    /// <param name="userPhotoSize">The desired size of the returned photo. Valid photo sizes are in UserPhotoSize</param>
    /// <param name="entityTag">A photo's cache ID which will allow the caller to ensure their cached photo is up to date</param>
    /// <param name="token"></param>
    /// <returns>A result object containing the photo state</returns>
    public async Task<GetUserPhotoResults> GetUserPhoto(
        string emailAddress,
        string userPhotoSize,
        string entityTag,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(emailAddress);
        EwsUtilities.ValidateParam(userPhotoSize);
        EwsUtilities.ValidateParamAllowNull(entityTag);

        var request = new GetUserPhotoRequest(this)
        {
            EmailAddress = emailAddress,
            UserPhotoSize = userPhotoSize,
            EntityTag = entityTag,
        };

        var response = await request.Execute(token).ConfigureAwait(false);
        return response.Results;
    }

    #endregion


    #region Attachment operations

    /// <summary>
    ///     Gets an attachment.
    /// </summary>
    /// <param name="attachments">The attachments.</param>
    /// <param name="bodyType">Type of the body.</param>
    /// <param name="additionalProperties">The additional properties.</param>
    /// <param name="errorHandling">Type of error handling to perform.</param>
    /// <param name="token"></param>
    /// <returns>Service response collection.</returns>
    private Task<ServiceResponseCollection<GetAttachmentResponse>> InternalGetAttachments(
        IEnumerable<Attachment> attachments,
        BodyType? bodyType,
        IEnumerable<PropertyDefinitionBase>? additionalProperties,
        ServiceErrorHandling errorHandling,
        CancellationToken token
    )
    {
        var request = new GetAttachmentRequest(this, errorHandling)
        {
            BodyType = bodyType,
        };

        request.Attachments.AddRange(attachments);

        if (additionalProperties != null)
        {
            request.AdditionalProperties.AddRange(additionalProperties);
        }

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Gets attachments.
    /// </summary>
    /// <param name="attachments">The attachments.</param>
    /// <param name="bodyType">Type of the body.</param>
    /// <param name="additionalProperties">The additional properties.</param>
    /// <param name="token"></param>
    /// <returns>Service response collection.</returns>
    public Task<ServiceResponseCollection<GetAttachmentResponse>> GetAttachments(
        Attachment[] attachments,
        BodyType? bodyType,
        IEnumerable<PropertyDefinitionBase> additionalProperties,
        CancellationToken token = default
    )
    {
        return InternalGetAttachments(
            attachments,
            bodyType,
            additionalProperties,
            ServiceErrorHandling.ReturnErrors,
            token
        );
    }

    /// <summary>
    ///     Gets attachments.
    /// </summary>
    /// <param name="attachmentIds">The attachment ids.</param>
    /// <param name="bodyType">Type of the body.</param>
    /// <param name="additionalProperties">The additional properties.</param>
    /// <param name="token"></param>
    /// <returns>Service response collection.</returns>
    public Task<ServiceResponseCollection<GetAttachmentResponse>> GetAttachments(
        string[] attachmentIds,
        BodyType? bodyType,
        IEnumerable<PropertyDefinitionBase>? additionalProperties,
        CancellationToken token = default
    )
    {
        var request = new GetAttachmentRequest(this, ServiceErrorHandling.ReturnErrors)
        {
            BodyType = bodyType,
        };

        request.AttachmentIds.AddRange(attachmentIds);

        if (additionalProperties != null)
        {
            request.AdditionalProperties.AddRange(additionalProperties);
        }

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Gets an attachment.
    /// </summary>
    /// <param name="attachment">The attachment.</param>
    /// <param name="bodyType">Type of the body.</param>
    /// <param name="additionalProperties">The additional properties.</param>
    /// <param name="token"></param>
    internal Task<ServiceResponseCollection<GetAttachmentResponse>> GetAttachment(
        Attachment attachment,
        BodyType? bodyType,
        IEnumerable<PropertyDefinitionBase>? additionalProperties,
        CancellationToken token
    )
    {
        return InternalGetAttachments(
            new[]
            {
                attachment,
            },
            bodyType,
            additionalProperties,
            ServiceErrorHandling.ThrowOnError,
            token
        );
    }

    /// <summary>
    ///     Creates attachments.
    /// </summary>
    /// <param name="parentItemId">The parent item id.</param>
    /// <param name="attachments">The attachments.</param>
    /// <param name="token"></param>
    /// <returns>Service response collection.</returns>
    internal Task<ServiceResponseCollection<CreateAttachmentResponse>> CreateAttachments(
        string parentItemId,
        IEnumerable<Attachment> attachments,
        CancellationToken token
    )
    {
        var request = new CreateAttachmentRequest(this, ServiceErrorHandling.ReturnErrors)
        {
            ParentItemId = parentItemId,
        };

        request.Attachments.AddRange(attachments);

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Deletes attachments.
    /// </summary>
    /// <param name="attachments">The attachments.</param>
    /// <param name="token"></param>
    /// <returns>Service response collection.</returns>
    internal Task<ServiceResponseCollection<DeleteAttachmentResponse>> DeleteAttachments(
        IEnumerable<Attachment> attachments,
        CancellationToken token
    )
    {
        var request = new DeleteAttachmentRequest(this, ServiceErrorHandling.ReturnErrors);

        request.Attachments.AddRange(attachments);

        return request.ExecuteAsync(token);
    }

    #endregion


    #region AD related operations

    /// <summary>
    ///     Finds contacts in the user's Contacts folder and the Global Address List (in that order) that have names
    ///     that match the one passed as a parameter. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="nameToResolve">The name to resolve.</param>
    /// <returns>A collection of name resolutions whose names match the one passed as a parameter.</returns>
    public Task<NameResolutionCollection> ResolveName(string nameToResolve)
    {
        return ResolveName(nameToResolve, ResolveNameSearchLocation.ContactsThenDirectory, false);
    }

    /// <summary>
    ///     Finds contacts in the Global Address List and/or in specific contact folders that have names
    ///     that match the one passed as a parameter. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="nameToResolve">The name to resolve.</param>
    /// <param name="parentFolderIds">The Ids of the contact folders in which to look for matching contacts.</param>
    /// <param name="searchScope">The scope of the search.</param>
    /// <param name="returnContactDetails">
    ///     Indicates whether full contact information should be returned for each of the found
    ///     contacts.
    /// </param>
    /// <returns>A collection of name resolutions whose names match the one passed as a parameter.</returns>
    public Task<NameResolutionCollection> ResolveName(
        string nameToResolve,
        IEnumerable<FolderId>? parentFolderIds,
        ResolveNameSearchLocation searchScope,
        bool returnContactDetails
    )
    {
        return ResolveName(nameToResolve, parentFolderIds, searchScope, returnContactDetails, null);
    }

    /// <summary>
    ///     Finds contacts in the Global Address List and/or in specific contact folders that have names
    ///     that match the one passed as a parameter. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="nameToResolve">The name to resolve.</param>
    /// <param name="parentFolderIds">The Ids of the contact folders in which to look for matching contacts.</param>
    /// <param name="searchScope">The scope of the search.</param>
    /// <param name="returnContactDetails">
    ///     Indicates whether full contact information should be returned for each of the found
    ///     contacts.
    /// </param>
    /// <param name="contactDataPropertySet">The property set for the contact details</param>
    /// <param name="token"></param>
    /// <returns>A collection of name resolutions whose names match the one passed as a parameter.</returns>
    public async Task<NameResolutionCollection> ResolveName(
        string nameToResolve,
        IEnumerable<FolderId>? parentFolderIds,
        ResolveNameSearchLocation searchScope,
        bool returnContactDetails,
        PropertySet? contactDataPropertySet,
        CancellationToken token = default
    )
    {
        if (contactDataPropertySet != null)
        {
            EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2010_SP1, "ResolveName");
        }

        EwsUtilities.ValidateParam(nameToResolve);
        if (parentFolderIds != null)
        {
            EwsUtilities.ValidateParamCollection(parentFolderIds);
        }

        var request = new ResolveNamesRequest(this)
        {
            NameToResolve = nameToResolve,
            ReturnFullContactData = returnContactDetails,
            SearchLocation = searchScope,
            ContactDataPropertySet = contactDataPropertySet,
        };

        request.ParentFolderIds.AddRange(parentFolderIds);

        var responses = await request.ExecuteAsync(token).ConfigureAwait(false);
        return responses[0].Resolutions;
    }

    /// <summary>
    ///     Finds contacts in the Global Address List that have names that match the one passed as a parameter.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="nameToResolve">The name to resolve.</param>
    /// <param name="searchScope">The scope of the search.</param>
    /// <param name="returnContactDetails">
    ///     Indicates whether full contact information should be returned for each of the found
    ///     contacts.
    /// </param>
    /// <param name="contactDataPropertySet">Property set for contact details</param>
    /// <returns>A collection of name resolutions whose names match the one passed as a parameter.</returns>
    public Task<NameResolutionCollection> ResolveName(
        string nameToResolve,
        ResolveNameSearchLocation searchScope,
        bool returnContactDetails,
        PropertySet contactDataPropertySet
    )
    {
        return ResolveName(nameToResolve, null, searchScope, returnContactDetails, contactDataPropertySet);
    }

    /// <summary>
    ///     Finds contacts in the Global Address List that have names that match the one passed as a parameter.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="nameToResolve">The name to resolve.</param>
    /// <param name="searchScope">The scope of the search.</param>
    /// <param name="returnContactDetails">
    ///     Indicates whether full contact information should be returned for each of the found
    ///     contacts.
    /// </param>
    /// <returns>A collection of name resolutions whose names match the one passed as a parameter.</returns>
    public Task<NameResolutionCollection> ResolveName(
        string nameToResolve,
        ResolveNameSearchLocation searchScope,
        bool returnContactDetails
    )
    {
        return ResolveName(nameToResolve, null, searchScope, returnContactDetails);
    }

    /// <summary>
    ///     Expands a group by retrieving a list of its members. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="emailAddress">The e-mail address of the group.</param>
    /// <param name="token"></param>
    /// <returns>An ExpandGroupResults containing the members of the group.</returns>
    public async Task<ExpandGroupResults> ExpandGroup(EmailAddress emailAddress, CancellationToken token = default)
    {
        EwsUtilities.ValidateParam(emailAddress);

        var request = new ExpandGroupRequest(this)
        {
            EmailAddress = emailAddress,
        };

        var responses = await request.ExecuteAsync(token).ConfigureAwait(false);
        return responses[0].Members;
    }

    /// <summary>
    ///     Expands a group by retrieving a list of its members. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="groupId">The Id of the group to expand.</param>
    /// <returns>An ExpandGroupResults containing the members of the group.</returns>
    public Task<ExpandGroupResults> ExpandGroup(ItemId groupId)
    {
        EwsUtilities.ValidateParam(groupId);

        var emailAddress = new EmailAddress
        {
            Id = groupId,
        };

        return ExpandGroup(emailAddress);
    }

    /// <summary>
    ///     Expands a group by retrieving a list of its members. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="smtpAddress">The SMTP address of the group to expand.</param>
    /// <returns>An ExpandGroupResults containing the members of the group.</returns>
    public Task<ExpandGroupResults> ExpandGroup(string smtpAddress)
    {
        EwsUtilities.ValidateParam(smtpAddress);

        return ExpandGroup(new EmailAddress(smtpAddress));
    }

    /// <summary>
    ///     Expands a group by retrieving a list of its members. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="address">The SMTP address of the group to expand.</param>
    /// <param name="routingType">The routing type of the address of the group to expand.</param>
    /// <returns>An ExpandGroupResults containing the members of the group.</returns>
    public Task<ExpandGroupResults> ExpandGroup(string address, string routingType)
    {
        EwsUtilities.ValidateParam(address);
        EwsUtilities.ValidateParam(routingType);

        var emailAddress = new EmailAddress(address)
        {
            RoutingType = routingType,
        };

        return ExpandGroup(emailAddress);
    }

    /// <summary>
    ///     Get the password expiration date
    /// </summary>
    /// <param name="mailboxSmtpAddress">The e-mail address of the user.</param>
    /// <param name="token"></param>
    /// <returns>The password expiration date.</returns>
    public async Task<DateTime?> GetPasswordExpirationDate(string mailboxSmtpAddress, CancellationToken token = default)
    {
        var request = new GetPasswordExpirationDateRequest(this)
        {
            MailboxSmtpAddress = mailboxSmtpAddress,
        };

        var response = await request.Execute(token).ConfigureAwait(false);
        return response.PasswordExpirationDate;
    }

    #endregion


    #region Notification operations

    /// <summary>
    ///     Subscribes to pull notifications. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
    /// <param name="timeout">
    ///     The timeout, in minutes, after which the subscription expires. Timeout must be between 1 and
    ///     1440.
    /// </param>
    /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
    /// <param name="token"></param>
    /// <param name="eventTypes">The event types to subscribe to.</param>
    /// <returns>A PullSubscription representing the new subscription.</returns>
    public async Task<PullSubscription> SubscribeToPullNotifications(
        IEnumerable<FolderId> folderIds,
        int timeout,
        string watermark,
        CancellationToken token = default,
        params EventType[] eventTypes
    )
    {
        EwsUtilities.ValidateParamCollection(folderIds);

        var responses = await BuildSubscribeToPullNotificationsRequest(folderIds, timeout, watermark, eventTypes)
            .ExecuteAsync(token)
            .ConfigureAwait(false);

        return responses[0].Subscription;
    }

    /// <summary>
    ///     Subscribes to pull notifications on all folders in the authenticated user's mailbox. Calling this method results in
    ///     a call to EWS.
    /// </summary>
    /// <param name="timeout">
    ///     The timeout, in minutes, after which the subscription expires. Timeout must be between 1 and
    ///     1440.
    /// </param>
    /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
    /// <param name="token"></param>
    /// <param name="eventTypes">The event types to subscribe to.</param>
    /// <returns>A PullSubscription representing the new subscription.</returns>
    public async Task<PullSubscription> SubscribeToPullNotificationsOnAllFolders(
        int timeout,
        string watermark,
        CancellationToken token = default,
        params EventType[] eventTypes
    )
    {
        EwsUtilities.ValidateMethodVersion(
            this,
            ExchangeVersion.Exchange2010,
            "SubscribeToPullNotificationsOnAllFolders"
        );

        var responses = await BuildSubscribeToPullNotificationsRequest(null, timeout, watermark, eventTypes)
            .ExecuteAsync(token)
            .ConfigureAwait(false);

        return responses[0].Subscription;
    }

    /// <summary>
    ///     Builds a request to subscribe to pull notifications in the authenticated user's mailbox.
    /// </summary>
    /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
    /// <param name="timeout">
    ///     The timeout, in minutes, after which the subscription expires. Timeout must be between 1 and
    ///     1440.
    /// </param>
    /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
    /// <param name="eventTypes">The event types to subscribe to.</param>
    /// <returns>A request to subscribe to pull notifications in the authenticated user's mailbox. </returns>
    private SubscribeToPullNotificationsRequest BuildSubscribeToPullNotificationsRequest(
        IEnumerable<FolderId>? folderIds,
        int timeout,
        string watermark,
        EventType[] eventTypes
    )
    {
        if (timeout < 1 || timeout > 1440)
        {
            throw new ArgumentOutOfRangeException(nameof(timeout), Strings.TimeoutMustBeBetween1And1440);
        }

        EwsUtilities.ValidateParamCollection(eventTypes);

        var request = new SubscribeToPullNotificationsRequest(this)
        {
            Timeout = timeout,
            Watermark = watermark,
        };

        if (folderIds != null)
        {
            request.FolderIds.AddRange(folderIds);
        }

        request.EventTypes.AddRange(eventTypes);

        return request;
    }

    /// <summary>
    ///     Unsubscribes from a subscription. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="subscriptionId">The Id of the pull subscription to unsubscribe from.</param>
    /// <param name="token"></param>
    internal System.Threading.Tasks.Task Unsubscribe(string subscriptionId, CancellationToken token)
    {
        return BuildUnsubscribeRequest(subscriptionId).ExecuteAsync(token);
    }

    /// <summary>
    ///     Builds a request to unsubscribe from a subscription.
    /// </summary>
    /// <param name="subscriptionId">The Id of the subscription for which to get the events.</param>
    /// <returns>A request to unsubscribe from a subscription.</returns>
    private UnsubscribeRequest BuildUnsubscribeRequest(string subscriptionId)
    {
        EwsUtilities.ValidateParam(subscriptionId);

        var request = new UnsubscribeRequest(this)
        {
            SubscriptionId = subscriptionId,
        };

        return request;
    }

    /// <summary>
    ///     Retrieves the latest events associated with a pull subscription. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="subscriptionId">The Id of the pull subscription for which to get the events.</param>
    /// <param name="watermark">The watermark representing the point in time where to start receiving events.</param>
    /// <param name="token"></param>
    /// <returns>A GetEventsResults containing a list of events associated with the subscription.</returns>
    internal async Task<GetEventsResults> GetEvents(string subscriptionId, string watermark, CancellationToken token)
    {
        var responses = await BuildGetEventsRequest(subscriptionId, watermark)
            .ExecuteAsync(token)
            .ConfigureAwait(false);

        return responses[0].Results;
    }

    /// <summary>
    ///     Builds an request to retrieve the latest events associated with a pull subscription.
    /// </summary>
    /// <param name="subscriptionId">The Id of the pull subscription for which to get the events.</param>
    /// <param name="watermark">The watermark representing the point in time where to start receiving events.</param>
    /// <returns>An request to retrieve the latest events associated with a pull subscription. </returns>
    private GetEventsRequest BuildGetEventsRequest(string subscriptionId, string watermark)
    {
        EwsUtilities.ValidateParam(subscriptionId);
        EwsUtilities.ValidateParam(watermark);

        var request = new GetEventsRequest(this)
        {
            SubscriptionId = subscriptionId,
            Watermark = watermark,
        };

        return request;
    }

    /// <summary>
    ///     Subscribes to push notifications. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
    /// <param name="url">The URL of the Web Service endpoint the Exchange server should push events to.</param>
    /// <param name="frequency">
    ///     The frequency, in minutes, at which the Exchange server should contact the Web Service
    ///     endpoint. Frequency must be between 1 and 1440.
    /// </param>
    /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
    /// <param name="token"></param>
    /// <param name="eventTypes">The event types to subscribe to.</param>
    /// <returns>A PushSubscription representing the new subscription.</returns>
    public async Task<PushSubscription> SubscribeToPushNotifications(
        IEnumerable<FolderId> folderIds,
        Uri url,
        int frequency,
        string watermark,
        CancellationToken token = default,
        params EventType[] eventTypes
    )
    {
        EwsUtilities.ValidateParamCollection(folderIds);

        var responses = await BuildSubscribeToPushNotificationsRequest(
                folderIds,
                url,
                frequency,
                watermark,
                null,
                null,
                eventTypes
            )
            .ExecuteAsync(token)
            .ConfigureAwait(false);

        return responses[0].Subscription;
    }

    /// <summary>
    ///     Subscribes to push notifications on all folders in the authenticated user's mailbox. Calling this method results in
    ///     a call to EWS.
    /// </summary>
    /// <param name="url">The URL of the Web Service endpoint the Exchange server should push events to.</param>
    /// <param name="frequency">
    ///     The frequency, in minutes, at which the Exchange server should contact the Web Service
    ///     endpoint. Frequency must be between 1 and 1440.
    /// </param>
    /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
    /// <param name="token"></param>
    /// <param name="eventTypes">The event types to subscribe to.</param>
    /// <returns>A PushSubscription representing the new subscription.</returns>
    public async Task<PushSubscription> SubscribeToPushNotificationsOnAllFolders(
        Uri url,
        int frequency,
        string watermark,
        CancellationToken token = default,
        params EventType[] eventTypes
    )
    {
        EwsUtilities.ValidateMethodVersion(
            this,
            ExchangeVersion.Exchange2010,
            "SubscribeToPushNotificationsOnAllFolders"
        );

        var responses = await BuildSubscribeToPushNotificationsRequest(
                null,
                url,
                frequency,
                watermark,
                null,
                null,
                eventTypes
            )
            .ExecuteAsync(token)
            .ConfigureAwait(false);

        return responses[0].Subscription;
    }

    /// <summary>
    ///     Subscribes to push notifications. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
    /// <param name="url">The URL of the Web Service endpoint the Exchange server should push events to.</param>
    /// <param name="frequency">
    ///     The frequency, in minutes, at which the Exchange server should contact the Web Service
    ///     endpoint. Frequency must be between 1 and 1440.
    /// </param>
    /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
    /// <param name="callerData">Optional caller data that will be returned the call back.</param>
    /// <param name="token"></param>
    /// <param name="eventTypes">The event types to subscribe to.</param>
    /// <returns>A PushSubscription representing the new subscription.</returns>
    public async Task<PushSubscription> SubscribeToPushNotifications(
        IEnumerable<FolderId> folderIds,
        Uri url,
        int frequency,
        string watermark,
        string callerData,
        CancellationToken token = default,
        params EventType[] eventTypes
    )
    {
        EwsUtilities.ValidateParamCollection(folderIds);

        var responses = await BuildSubscribeToPushNotificationsRequest(
                folderIds,
                url,
                frequency,
                watermark,
                callerData,
                null, // AnchorMailbox
                eventTypes
            )
            .ExecuteAsync(token)
            .ConfigureAwait(false);
        return responses[0].Subscription;
    }

    /// <summary>
    ///     Subscribes to push notifications on a group mailbox. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="groupMailboxSmtp">The smtpaddress of the group mailbox to subscribe to.</param>
    /// <param name="url">The URL of the Web Service endpoint the Exchange server should push events to.</param>
    /// <param name="frequency">
    ///     The frequency, in minutes, at which the Exchange server should contact the Web Service
    ///     endpoint. Frequency must be between 1 and 1440.
    /// </param>
    /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
    /// <param name="callerData">Optional caller data that will be returned the call back.</param>
    /// <param name="token"></param>
    /// <param name="eventTypes">The event types to subscribe to.</param>
    /// <returns>A PushSubscription representing the new subscription.</returns>
    public async Task<PushSubscription> SubscribeToGroupPushNotifications(
        string groupMailboxSmtp,
        Uri url,
        int frequency,
        string watermark,
        string callerData,
        CancellationToken token = default,
        params EventType[] eventTypes
    )
    {
        var folderIds = new[]
        {
            new FolderId(WellKnownFolderName.Inbox, new Mailbox(groupMailboxSmtp)),
        };

        var responses = await BuildSubscribeToPushNotificationsRequest(
                folderIds,
                url,
                frequency,
                watermark,
                callerData,
                groupMailboxSmtp, // AnchorMailbox
                eventTypes
            )
            .ExecuteAsync(token)
            .ConfigureAwait(false);

        return responses[0].Subscription;
    }

    /// <summary>
    ///     Subscribes to push notifications on all folders in the authenticated user's mailbox. Calling this method results in
    ///     a call to EWS.
    /// </summary>
    /// <param name="url">The URL of the Web Service endpoint the Exchange server should push events to.</param>
    /// <param name="frequency">
    ///     The frequency, in minutes, at which the Exchange server should contact the Web Service
    ///     endpoint. Frequency must be between 1 and 1440.
    /// </param>
    /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
    /// <param name="callerData">Optional caller data that will be returned the call back.</param>
    /// <param name="token"></param>
    /// <param name="eventTypes">The event types to subscribe to.</param>
    /// <returns>A PushSubscription representing the new subscription.</returns>
    public async Task<PushSubscription> SubscribeToPushNotificationsOnAllFolders(
        Uri url,
        int frequency,
        string watermark,
        string callerData,
        CancellationToken token = default,
        params EventType[] eventTypes
    )
    {
        EwsUtilities.ValidateMethodVersion(
            this,
            ExchangeVersion.Exchange2010,
            "SubscribeToPushNotificationsOnAllFolders"
        );

        var responses = await BuildSubscribeToPushNotificationsRequest(
                null,
                url,
                frequency,
                watermark,
                callerData,
                null, // AnchorMailbox
                eventTypes
            )
            .ExecuteAsync(token)
            .ConfigureAwait(false);

        return responses[0].Subscription;
    }

    /// <summary>
    ///     Set a TeamMailbox
    /// </summary>
    /// <param name="emailAddress">TeamMailbox email address</param>
    /// <param name="sharePointSiteUrl">SharePoint site URL</param>
    /// <param name="state">TeamMailbox lifecycle state</param>
    /// <param name="token"></param>
    public System.Threading.Tasks.Task SetTeamMailbox(
        EmailAddress emailAddress,
        Uri sharePointSiteUrl,
        TeamMailboxLifecycleState state,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013, "SetTeamMailbox");

        if (emailAddress == null)
        {
            throw new ArgumentNullException(nameof(emailAddress));
        }

        if (sharePointSiteUrl == null)
        {
            throw new ArgumentNullException(nameof(sharePointSiteUrl));
        }

        var request = new SetTeamMailboxRequest(this, emailAddress, sharePointSiteUrl, state);
        return request.Execute(token);
    }

    /// <summary>
    ///     Unpin a TeamMailbox
    /// </summary>
    /// <param name="emailAddress">TeamMailbox email address</param>
    /// <param name="token"></param>
    public System.Threading.Tasks.Task UnpinTeamMailbox(EmailAddress emailAddress, CancellationToken token = default)
    {
        EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013, "UnpinTeamMailbox");

        if (emailAddress == null)
        {
            throw new ArgumentNullException(nameof(emailAddress));
        }

        var request = new UnpinTeamMailboxRequest(this, emailAddress);
        return request.Execute(token);
    }

    /// <summary>
    ///     Builds an request to request to subscribe to push notifications in the authenticated user's mailbox.
    /// </summary>
    /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
    /// <param name="url">The URL of the Web Service endpoint the Exchange server should push events to.</param>
    /// <param name="frequency">
    ///     The frequency, in minutes, at which the Exchange server should contact the Web Service
    ///     endpoint. Frequency must be between 1 and 1440.
    /// </param>
    /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
    /// <param name="callerData">Optional caller data that will be returned the call back.</param>
    /// <param name="anchorMailbox">The smtpaddress of the mailbox to subscribe to.</param>
    /// <param name="eventTypes">The event types to subscribe to.</param>
    /// <returns>A request to request to subscribe to push notifications in the authenticated user's mailbox.</returns>
    private SubscribeToPushNotificationsRequest BuildSubscribeToPushNotificationsRequest(
        IEnumerable<FolderId>? folderIds,
        Uri url,
        int frequency,
        string watermark,
        string? callerData,
        string? anchorMailbox,
        EventType[] eventTypes
    )
    {
        EwsUtilities.ValidateParam(url);

        if (frequency < 1 || frequency > 1440)
        {
            throw new ArgumentOutOfRangeException(nameof(frequency), Strings.FrequencyMustBeBetween1And1440);
        }

        EwsUtilities.ValidateParamCollection(eventTypes);

        var request = new SubscribeToPushNotificationsRequest(this)
        {
            AnchorMailbox = anchorMailbox,
            Url = url,
            Frequency = frequency,
            Watermark = watermark,
            CallerData = callerData,
        };

        if (folderIds != null)
        {
            request.FolderIds.AddRange(folderIds);
        }

        request.EventTypes.AddRange(eventTypes);

        return request;
    }

    /// <summary>
    ///     Subscribes to streaming notifications. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
    /// <param name="eventTypes">The event types to subscribe to.</param>
    /// <param name="token"></param>
    /// <returns>A StreamingSubscription representing the new subscription.</returns>
    public async Task<StreamingSubscription> SubscribeToStreamingNotifications(
        IEnumerable<FolderId> folderIds,
        EventType[] eventTypes,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2010_SP1, "SubscribeToStreamingNotifications");

        EwsUtilities.ValidateParamCollection(folderIds);

        var responses = await BuildSubscribeToStreamingNotificationsRequest(folderIds, eventTypes)
            .ExecuteAsync(token)
            .ConfigureAwait(false);

        return responses[0].Subscription;
    }

    /// <summary>
    ///     Subscribes to streaming notifications on all folders in the authenticated user's mailbox. Calling this method
    ///     results in a call to EWS.
    /// </summary>
    /// <param name="token"></param>
    /// <param name="eventTypes">The event types to subscribe to.</param>
    /// <returns>A StreamingSubscription representing the new subscription.</returns>
    public async Task<StreamingSubscription> SubscribeToStreamingNotificationsOnAllFolders(
        CancellationToken token = default,
        params EventType[] eventTypes
    )
    {
        EwsUtilities.ValidateMethodVersion(
            this,
            ExchangeVersion.Exchange2010_SP1,
            "SubscribeToStreamingNotificationsOnAllFolders"
        );

        var responses = await BuildSubscribeToStreamingNotificationsRequest(null, eventTypes)
            .ExecuteAsync(token)
            .ConfigureAwait(false);

        return responses[0].Subscription;
    }

    /// <summary>
    ///     Builds request to subscribe to streaming notifications in the authenticated user's mailbox.
    /// </summary>
    /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
    /// <param name="eventTypes">The event types to subscribe to.</param>
    /// <returns>A request to subscribe to streaming notifications in the authenticated user's mailbox. </returns>
    private SubscribeToStreamingNotificationsRequest BuildSubscribeToStreamingNotificationsRequest(
        IEnumerable<FolderId>? folderIds,
        EventType[] eventTypes
    )
    {
        EwsUtilities.ValidateParamCollection(eventTypes);

        var request = new SubscribeToStreamingNotificationsRequest(this);

        if (folderIds != null)
        {
            request.FolderIds.AddRange(folderIds);
        }

        request.EventTypes.AddRange(eventTypes);

        return request;
    }

    #endregion


    #region Synchronization operations

    /// <summary>
    ///     Synchronizes the items of a specific folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="syncFolderId">The Id of the folder containing the items to synchronize with.</param>
    /// <param name="propertySet">The set of properties to retrieve for synchronized items.</param>
    /// <param name="ignoredItemIds">The optional list of item Ids that should be ignored.</param>
    /// <param name="maxChangesReturned">The maximum number of changes that should be returned.</param>
    /// <param name="syncScope">The sync scope identifying items to include in the ChangeCollection.</param>
    /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
    /// <returns>A ChangeCollection containing a list of changes that occurred in the specified folder.</returns>
    public Task<ChangeCollection<ItemChange>> SyncFolderItems(
        FolderId syncFolderId,
        PropertySet propertySet,
        IEnumerable<ItemId> ignoredItemIds,
        int maxChangesReturned,
        SyncFolderItemsScope syncScope,
        string syncState
    )
    {
        return SyncFolderItems(syncFolderId, propertySet, ignoredItemIds, maxChangesReturned, 0, syncScope, syncState);
    }

    /// <summary>
    ///     Synchronizes the items of a specific folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="syncFolderId">The Id of the folder containing the items to synchronize with.</param>
    /// <param name="propertySet">The set of properties to retrieve for synchronized items.</param>
    /// <param name="ignoredItemIds">The optional list of item Ids that should be ignored.</param>
    /// <param name="maxChangesReturned">The maximum number of changes that should be returned.</param>
    /// <param name="numberOfDays">Limit the changes returned to this many days ago; 0 means no limit.</param>
    /// <param name="syncScope">The sync scope identifying items to include in the ChangeCollection.</param>
    /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
    /// <param name="token"></param>
    /// <returns>A ChangeCollection containing a list of changes that occurred in the specified folder.</returns>
    public async Task<ChangeCollection<ItemChange>> SyncFolderItems(
        FolderId syncFolderId,
        PropertySet propertySet,
        IEnumerable<ItemId> ignoredItemIds,
        int maxChangesReturned,
        int numberOfDays,
        SyncFolderItemsScope syncScope,
        string syncState,
        CancellationToken token = default
    )
    {
        var responses = await BuildSyncFolderItemsRequest(
                syncFolderId,
                propertySet,
                ignoredItemIds,
                maxChangesReturned,
                numberOfDays,
                syncScope,
                syncState
            )
            .ExecuteAsync(token)
            .ConfigureAwait(false);

        return responses[0].Changes;
    }

    /// <summary>
    ///     Builds a request to synchronize the items of a specific folder.
    /// </summary>
    /// <param name="syncFolderId">The Id of the folder containing the items to synchronize with.</param>
    /// <param name="propertySet">The set of properties to retrieve for synchronized items.</param>
    /// <param name="ignoredItemIds">The optional list of item Ids that should be ignored.</param>
    /// <param name="maxChangesReturned">The maximum number of changes that should be returned.</param>
    /// <param name="numberOfDays">Limit the changes returned to this many days ago; 0 means no limit.</param>
    /// <param name="syncScope">The sync scope identifying items to include in the ChangeCollection.</param>
    /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
    /// <returns>A request to synchronize the items of a specific folder.</returns>
    private SyncFolderItemsRequest BuildSyncFolderItemsRequest(
        FolderId syncFolderId,
        PropertySet propertySet,
        IEnumerable<ItemId>? ignoredItemIds,
        int maxChangesReturned,
        int numberOfDays,
        SyncFolderItemsScope syncScope,
        string syncState
    )
    {
        EwsUtilities.ValidateParam(syncFolderId);
        EwsUtilities.ValidateParam(propertySet);

        var request = new SyncFolderItemsRequest(this)
        {
            SyncFolderId = syncFolderId,
            PropertySet = propertySet,
            MaxChangesReturned = maxChangesReturned,
            NumberOfDays = numberOfDays,
            SyncScope = syncScope,
            SyncState = syncState,
        };

        if (ignoredItemIds != null)
        {
            request.IgnoredItemIds.AddRange(ignoredItemIds);
        }

        return request;
    }

    /// <summary>
    ///     Synchronizes the sub-folders of a specific folder. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="syncFolderId">
    ///     The Id of the folder containing the items to synchronize with. A null value indicates the
    ///     root folder of the mailbox.
    /// </param>
    /// <param name="propertySet">The set of properties to retrieve for synchronized items.</param>
    /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
    /// <param name="token"></param>
    /// <returns>A ChangeCollection containing a list of changes that occurred in the specified folder.</returns>
    public async Task<ChangeCollection<FolderChange>> SyncFolderHierarchy(
        FolderId? syncFolderId,
        PropertySet propertySet,
        string syncState,
        CancellationToken token = default
    )
    {
        var responses = await BuildSyncFolderHierarchyRequest(syncFolderId, propertySet, syncState)
            .ExecuteAsync(token)
            .ConfigureAwait(false);

        return responses[0].Changes;
    }

    /// <summary>
    ///     Synchronizes the entire folder hierarchy of the mailbox this Service is connected to. Calling this method results
    ///     in a call to EWS.
    /// </summary>
    /// <param name="propertySet">The set of properties to retrieve for synchronized items.</param>
    /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
    /// <returns>A ChangeCollection containing a list of changes that occurred in the specified folder.</returns>
    public Task<ChangeCollection<FolderChange>> SyncFolderHierarchy(PropertySet propertySet, string syncState)
    {
        return SyncFolderHierarchy(null, propertySet, syncState);
    }

    /// <summary>
    ///     Builds a request to synchronize the specified folder hierarchy of the mailbox this Service is connected to.
    /// </summary>
    /// <param name="syncFolderId">
    ///     The Id of the folder containing the items to synchronize with. A null value indicates the
    ///     root folder of the mailbox.
    /// </param>
    /// <param name="propertySet">The set of properties to retrieve for synchronized items.</param>
    /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
    /// <returns>A request to synchronize the specified folder hierarchy of the mailbox this Service is connected to.</returns>
    private SyncFolderHierarchyRequest BuildSyncFolderHierarchyRequest(
        FolderId syncFolderId,
        PropertySet propertySet,
        string syncState
    )
    {
        EwsUtilities.ValidateParamAllowNull(syncFolderId); // Null syncFolderId is allowed
        EwsUtilities.ValidateParam(propertySet);

        var request = new SyncFolderHierarchyRequest(this)
        {
            PropertySet = propertySet,
            SyncFolderId = syncFolderId,
            SyncState = syncState,
        };

        return request;
    }

    #endregion


    #region Availability operations

    /// <summary>
    ///     Gets Out of Office (OOF) settings for a specific user. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="smtpAddress">The SMTP address of the user for which to retrieve OOF settings.</param>
    /// <param name="token"></param>
    /// <returns>An OofSettings instance containing OOF information for the specified user.</returns>
    public async Task<OofSettings> GetUserOofSettings(string smtpAddress, CancellationToken token = default)
    {
        EwsUtilities.ValidateParam(smtpAddress);

        var request = new GetUserOofSettingsRequest(this)
        {
            SmtpAddress = smtpAddress,
        };

        var response = await request.Execute(token).ConfigureAwait(false);
        return response.OofSettings;
    }

    /// <summary>
    ///     Sets the Out of Office (OOF) settings for a specific mailbox. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="smtpAddress">The SMTP address of the user for which to set OOF settings.</param>
    /// <param name="oofSettings">The OOF settings.</param>
    /// <param name="token"></param>
    public System.Threading.Tasks.Task SetUserOofSettings(
        string smtpAddress,
        OofSettings oofSettings,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(smtpAddress);
        EwsUtilities.ValidateParam(oofSettings);

        var request = new SetUserOofSettingsRequest(this)
        {
            SmtpAddress = smtpAddress,
            OofSettings = oofSettings,
        };

        return request.Execute(token);
    }

    /// <summary>
    ///     Gets detailed information about the availability of a set of users, rooms, and resources within a
    ///     specified time window.
    /// </summary>
    /// <param name="attendees">The attendees for which to retrieve availability information.</param>
    /// <param name="timeWindow">The time window in which to retrieve user availability information.</param>
    /// <param name="requestedData">The requested data (free/busy and/or suggestions).</param>
    /// <param name="options">The options controlling the information returned.</param>
    /// <param name="token"></param>
    /// <returns>
    ///     The availability information for each user appears in a unique FreeBusyResponse object. The order of users
    ///     in the request determines the order of availability data for each user in the response.
    /// </returns>
    public Task<GetUserAvailabilityResults> GetUserAvailability(
        IEnumerable<AttendeeInfo> attendees,
        TimeWindow timeWindow,
        AvailabilityData requestedData,
        AvailabilityOptions options,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParamCollection(attendees);
        EwsUtilities.ValidateParam(timeWindow);
        EwsUtilities.ValidateParam(options);

        var request = new GetUserAvailabilityRequest(this)
        {
            Attendees = attendees,
            TimeWindow = timeWindow,
            RequestedData = requestedData,
            Options = options,
        };

        return request.Execute(token);
    }

    /// <summary>
    ///     Gets detailed information about the availability of a set of users, rooms, and resources within a
    ///     specified time window.
    /// </summary>
    /// <param name="attendees">The attendees for which to retrieve availability information.</param>
    /// <param name="timeWindow">The time window in which to retrieve user availability information.</param>
    /// <param name="requestedData">The requested data (free/busy and/or suggestions).</param>
    /// <returns>
    ///     The availability information for each user appears in a unique FreeBusyResponse object. The order of users
    ///     in the request determines the order of availability data for each user in the response.
    /// </returns>
    public Task<GetUserAvailabilityResults> GetUserAvailability(
        IEnumerable<AttendeeInfo> attendees,
        TimeWindow timeWindow,
        AvailabilityData requestedData
    )
    {
        return GetUserAvailability(attendees, timeWindow, requestedData, new AvailabilityOptions());
    }

    /// <summary>
    ///     Retrieves a collection of all room lists in the organization.
    /// </summary>
    /// <returns>An EmailAddressCollection containing all the room lists in the organization.</returns>
    public async Task<EmailAddressCollection> GetRoomLists(CancellationToken token = default)
    {
        var request = new GetRoomListsRequest(this);

        var response = await request.Execute(token).ConfigureAwait(false);
        return response.RoomLists;
    }

    /// <summary>
    ///     Retrieves a collection of all rooms in the specified room list in the organization.
    /// </summary>
    /// <param name="emailAddress">The e-mail address of the room list.</param>
    /// <param name="token"></param>
    /// <returns>A collection of EmailAddress objects representing all the rooms within the specified room list.</returns>
    public async Task<Collection<EmailAddress>> GetRooms(EmailAddress emailAddress, CancellationToken token = default)
    {
        EwsUtilities.ValidateParam(emailAddress);

        var request = new GetRoomsRequest(this)
        {
            RoomList = emailAddress,
        };

        var response = await request.Execute(token).ConfigureAwait(false);
        return response.Rooms;
    }

    #endregion


    #region Conversation

    /// <summary>
    ///     Retrieves a collection of all Conversations in the specified Folder.
    /// </summary>
    /// <param name="view">The view controlling the number of conversations returned.</param>
    /// <param name="folderId">The Id of the folder in which to search for conversations.</param>
    /// <param name="token"></param>
    /// <returns>Collection of conversations.</returns>
    public async Task<ICollection<Conversation>> FindConversation(
        ViewBase view,
        FolderId folderId,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(view);
        EwsUtilities.ValidateParam(folderId);
        EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2010_SP1, "FindConversation");

        var request = new FindConversationRequest(this)
        {
            View = view,
            FolderId = new FolderIdWrapper(folderId),
        };

        var response = await request.Execute(token).ConfigureAwait(false);
        return response.Conversations;
    }

    /// <summary>
    ///     Retrieves a collection of all Conversations in the specified Folder.
    /// </summary>
    /// <param name="view">The view controlling the number of conversations returned.</param>
    /// <param name="folderId">The Id of the folder in which to search for conversations.</param>
    /// <param name="anchorMailbox">The anchorMailbox Smtp address to route the request directly to group mailbox.</param>
    /// <param name="token"></param>
    /// <returns>Collection of conversations.</returns>
    /// <remarks>
    ///     This API designed to be used primarily in groups scenarios where we want to set the
    ///     anchor mailbox header so that request is routed directly to the group mailbox backend server.
    /// </remarks>
    public async Task<Collection<Conversation>> FindGroupConversation(
        ViewBase view,
        FolderId folderId,
        string anchorMailbox,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(view);
        EwsUtilities.ValidateParam(folderId);
        EwsUtilities.ValidateParam(anchorMailbox);
        EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2015, "FindConversation");

        var request = new FindConversationRequest(this)
        {
            View = view,
            FolderId = new FolderIdWrapper(folderId),
            AnchorMailbox = anchorMailbox,
        };

        var response = await request.Execute(token).ConfigureAwait(false);
        return response.Conversations;
    }

    /// <summary>
    ///     Retrieves a collection of all Conversations in the specified Folder.
    /// </summary>
    /// <param name="view">The view controlling the number of conversations returned.</param>
    /// <param name="folderId">The Id of the folder in which to search for conversations.</param>
    /// <param name="queryString">The query string for which the search is being performed</param>
    /// <param name="token"></param>
    /// <returns>Collection of conversations.</returns>
    public async Task<ICollection<Conversation>> FindConversation(
        ViewBase view,
        FolderId folderId,
        string queryString,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(view);
        EwsUtilities.ValidateParamAllowNull(queryString);
        EwsUtilities.ValidateParam(folderId);
        EwsUtilities.ValidateMethodVersion(
            this,
            ExchangeVersion.Exchange2013, // This method is only applicable for Exchange2013
            "FindConversation"
        );

        var request = new FindConversationRequest(this)
        {
            View = view,
            FolderId = new FolderIdWrapper(folderId),
            QueryString = queryString,
        };

        var response = await request.Execute(token).ConfigureAwait(false);
        return response.Conversations;
    }

    /// <summary>
    ///     Searches for and retrieves a collection of Conversations in the specified Folder.
    ///     Along with conversations, a list of highlight terms are returned.
    /// </summary>
    /// <param name="view">The view controlling the number of conversations returned.</param>
    /// <param name="folderId">The Id of the folder in which to search for conversations.</param>
    /// <param name="queryString">The query string for which the search is being performed</param>
    /// <param name="returnHighlightTerms">Flag indicating if highlight terms should be returned in the response</param>
    /// <param name="token"></param>
    /// <returns>FindConversation results.</returns>
    public async Task<FindConversationResults> FindConversation(
        ViewBase view,
        FolderId folderId,
        string queryString,
        bool returnHighlightTerms,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(view);
        EwsUtilities.ValidateParamAllowNull(queryString);
        EwsUtilities.ValidateParam(returnHighlightTerms);
        EwsUtilities.ValidateParam(folderId);
        EwsUtilities.ValidateMethodVersion(
            this,
            ExchangeVersion.Exchange2013, // This method is only applicable for Exchange2013
            "FindConversation"
        );

        var request = new FindConversationRequest(this)
        {
            View = view,
            FolderId = new FolderIdWrapper(folderId),
            QueryString = queryString,
            ReturnHighlightTerms = returnHighlightTerms,
        };

        var response = await request.Execute(token).ConfigureAwait(false);
        return response.Results;
    }

    /// <summary>
    ///     Searches for and retrieves a collection of Conversations in the specified Folder.
    ///     Along with conversations, a list of highlight terms are returned.
    /// </summary>
    /// <param name="view">The view controlling the number of conversations returned.</param>
    /// <param name="folderId">The Id of the folder in which to search for conversations.</param>
    /// <param name="queryString">The query string for which the search is being performed</param>
    /// <param name="returnHighlightTerms">Flag indicating if highlight terms should be returned in the response</param>
    /// <param name="mailboxScope">The mailbox scope to reference.</param>
    /// <param name="token"></param>
    /// <returns>FindConversation results.</returns>
    public async Task<FindConversationResults> FindConversation(
        ViewBase view,
        FolderId folderId,
        string queryString,
        bool returnHighlightTerms,
        MailboxSearchLocation? mailboxScope,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(view);
        EwsUtilities.ValidateParamAllowNull(queryString);
        EwsUtilities.ValidateParam(returnHighlightTerms);
        EwsUtilities.ValidateParam(folderId);

        EwsUtilities.ValidateMethodVersion(
            this,
            ExchangeVersion.Exchange2013, // This method is only applicable for Exchange2013
            "FindConversation"
        );

        var request = new FindConversationRequest(this)
        {
            View = view,
            FolderId = new FolderIdWrapper(folderId),
            QueryString = queryString,
            ReturnHighlightTerms = returnHighlightTerms,
            MailboxScope = mailboxScope,
        };

        var response = await request.Execute(token).ConfigureAwait(false);
        return response.Results;
    }

    /// <summary>
    ///     Gets the items for a set of conversations.
    /// </summary>
    /// <param name="conversations">Conversations with items to load.</param>
    /// <param name="propertySet">The set of properties to load.</param>
    /// <param name="foldersToIgnore">The folders to ignore.</param>
    /// <param name="sortOrder">Sort order of conversation tree nodes.</param>
    /// <param name="mailboxScope">The mailbox scope to reference.</param>
    /// <param name="maxItemsToReturn">Maximum number of items to return.</param>
    /// <param name="anchorMailbox">The smtpaddress of the mailbox that hosts the conversations</param>
    /// <param name="errorHandling">What type of error handling should be performed.</param>
    /// <param name="token"></param>
    /// <returns>GetConversationItems response.</returns>
    internal Task<ServiceResponseCollection<GetConversationItemsResponse>> InternalGetConversationItems(
        IEnumerable<ConversationRequest> conversations,
        PropertySet propertySet,
        IEnumerable<FolderId> foldersToIgnore,
        ConversationSortOrder? sortOrder,
        MailboxSearchLocation? mailboxScope,
        int? maxItemsToReturn,
        string? anchorMailbox,
        ServiceErrorHandling errorHandling,
        CancellationToken token
    )
    {
        EwsUtilities.ValidateParam(conversations);
        EwsUtilities.ValidateParam(propertySet, "itemProperties");
        EwsUtilities.ValidateParamAllowNull(foldersToIgnore);
        EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013, "GetConversationItems");

        var request = new GetConversationItemsRequest(this, errorHandling)
        {
            ItemProperties = propertySet,
            FoldersToIgnore = new FolderIdCollection(foldersToIgnore),
            SortOrder = sortOrder,
            MailboxScope = mailboxScope,
            MaxItemsToReturn = maxItemsToReturn,
            AnchorMailbox = anchorMailbox,
            Conversations = conversations.ToList(),
        };

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Gets the items for a set of conversations.
    /// </summary>
    /// <param name="conversations">Conversations with items to load.</param>
    /// <param name="propertySet">The set of properties to load.</param>
    /// <param name="foldersToIgnore">The folders to ignore.</param>
    /// <param name="sortOrder">Conversation item sort order.</param>
    /// <param name="token"></param>
    /// <returns>GetConversationItems response.</returns>
    public Task<ServiceResponseCollection<GetConversationItemsResponse>> GetConversationItems(
        IEnumerable<ConversationRequest> conversations,
        PropertySet propertySet,
        IEnumerable<FolderId> foldersToIgnore,
        ConversationSortOrder? sortOrder,
        CancellationToken token = default
    )
    {
        return InternalGetConversationItems(
            conversations,
            propertySet,
            foldersToIgnore,
            null, /* sortOrder */
            null, /* mailboxScope */
            null, /* maxItemsToReturn*/
            null, /* anchorMailbox */
            ServiceErrorHandling.ReturnErrors,
            token
        );
    }

    /// <summary>
    ///     Gets the items for a conversation.
    /// </summary>
    /// <param name="conversationId">The conversation id.</param>
    /// <param name="propertySet">The set of properties to load.</param>
    /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
    /// <param name="foldersToIgnore">The folders to ignore.</param>
    /// <param name="sortOrder">Conversation item sort order.</param>
    /// <param name="token"></param>
    /// <returns>ConversationResponseType response.</returns>
    public async Task<ConversationResponse> GetConversationItems(
        ConversationId conversationId,
        PropertySet propertySet,
        string syncState,
        IEnumerable<FolderId> foldersToIgnore,
        ConversationSortOrder? sortOrder,
        CancellationToken token = default
    )
    {
        var conversations = new List<ConversationRequest>
        {
            new(conversationId, syncState),
        };

        var responses = await InternalGetConversationItems(
                conversations,
                propertySet,
                foldersToIgnore,
                sortOrder,
                null,
                null,
                null,
                ServiceErrorHandling.ThrowOnError,
                token
            )
            .ConfigureAwait(false);
        return responses[0].Conversation;
    }

    /// <summary>
    ///     Gets the items for a conversation.
    /// </summary>
    /// <param name="conversationId">The conversation id.</param>
    /// <param name="propertySet">The set of properties to load.</param>
    /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
    /// <param name="foldersToIgnore">The folders to ignore.</param>
    /// <param name="sortOrder">Conversation item sort order.</param>
    /// <param name="anchorMailbox">The smtp address of the mailbox hosting the conversations</param>
    /// <param name="token"></param>
    /// <returns>ConversationResponseType response.</returns>
    /// <remarks>
    ///     This API designed to be used primarily in groups scenarios where we want to set the
    ///     anchor mailbox header so that request is routed directly to the group mailbox backend server.
    /// </remarks>
    public async Task<ConversationResponse> GetGroupConversationItems(
        ConversationId conversationId,
        PropertySet propertySet,
        string syncState,
        IEnumerable<FolderId> foldersToIgnore,
        ConversationSortOrder? sortOrder,
        string anchorMailbox,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(anchorMailbox);

        var conversations = new List<ConversationRequest>
        {
            new(conversationId, syncState),
        };

        var responses = await InternalGetConversationItems(
                conversations,
                propertySet,
                foldersToIgnore,
                sortOrder,
                null,
                null,
                anchorMailbox,
                ServiceErrorHandling.ThrowOnError,
                token
            )
            .ConfigureAwait(false);

        return responses[0].Conversation;
    }

    /// <summary>
    ///     Gets the items for a set of conversations.
    /// </summary>
    /// <param name="conversations">Conversations with items to load.</param>
    /// <param name="propertySet">The set of properties to load.</param>
    /// <param name="foldersToIgnore">The folders to ignore.</param>
    /// <param name="sortOrder">Conversation item sort order.</param>
    /// <param name="mailboxScope">The mailbox scope to reference.</param>
    /// <param name="token"></param>
    /// <returns>GetConversationItems response.</returns>
    public Task<ServiceResponseCollection<GetConversationItemsResponse>> GetConversationItems(
        IEnumerable<ConversationRequest> conversations,
        PropertySet propertySet,
        IEnumerable<FolderId> foldersToIgnore,
        ConversationSortOrder? sortOrder,
        MailboxSearchLocation? mailboxScope,
        CancellationToken token = default
    )
    {
        return InternalGetConversationItems(
            conversations,
            propertySet,
            foldersToIgnore,
            null,
            mailboxScope,
            null,
            null,
            ServiceErrorHandling.ReturnErrors,
            token
        );
    }

    /// <summary>
    ///     Applies ConversationAction on the specified conversation.
    /// </summary>
    /// <param name="actionType">ConversationAction</param>
    /// <param name="conversationIds">The conversation ids.</param>
    /// <param name="processRightAway">
    ///     True to process at once . This is blocking
    ///     and false to let the Assistant process it in the back ground
    /// </param>
    /// <param name="categories">Categories that need to be stamped can be null or empty</param>
    /// <param name="enableAlwaysDelete">
    ///     True moves every current and future messages in the conversation
    ///     to deleted items folder. False stops the always delete action. This is applicable only if
    ///     the action is AlwaysDelete
    /// </param>
    /// <param name="destinationFolderId">
    ///     Applicable if the action is AlwaysMove. This moves every current message and future
    ///     message in the conversation to the specified folder. Can be null if tis is then it stops
    ///     the always move action
    /// </param>
    /// <param name="errorHandlingMode">The error handling mode.</param>
    /// <param name="token"></param>
    /// <returns></returns>
    private Task<ServiceResponseCollection<ServiceResponse>> ApplyConversationAction(
        ConversationActionType actionType,
        IEnumerable<ConversationId> conversationIds,
        bool processRightAway,
        StringList? categories,
        bool enableAlwaysDelete,
        FolderId? destinationFolderId,
        ServiceErrorHandling errorHandlingMode,
        CancellationToken token
    )
    {
        EwsUtilities.Assert(
            actionType == ConversationActionType.AlwaysCategorize ||
            actionType == ConversationActionType.AlwaysMove ||
            actionType == ConversationActionType.AlwaysDelete,
            "ApplyConversationAction",
            "Invalid actionType"
        );

        EwsUtilities.ValidateParam(conversationIds, "conversationId");
        EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2010_SP1, "ApplyConversationAction");

        var request = new ApplyConversationActionRequest(this, errorHandlingMode);

        foreach (var conversationId in conversationIds)
        {
            var action = new ConversationAction
            {
                Action = actionType,
                ConversationId = conversationId,
                ProcessRightAway = processRightAway,
                Categories = categories,
                EnableAlwaysDelete = enableAlwaysDelete,
                DestinationFolderId = destinationFolderId != null ? new FolderIdWrapper(destinationFolderId) : null,
            };

            request.ConversationActions.Add(action);
        }

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Applies one time conversation action on items in specified folder inside
    ///     the conversation.
    /// </summary>
    /// <param name="actionType">The action.</param>
    /// <param name="idTimePairs">The id time pairs.</param>
    /// <param name="contextFolderId">The context folder id.</param>
    /// <param name="destinationFolderId">The destination folder id.</param>
    /// <param name="deleteType">Type of the delete.</param>
    /// <param name="isRead">The is read.</param>
    /// <param name="retentionPolicyType">Retention policy type.</param>
    /// <param name="retentionPolicyTagId">Retention policy tag id.  Null will clear the policy.</param>
    /// <param name="flag">Flag status.</param>
    /// <param name="suppressReadReceipts">Suppress read receipts flag.</param>
    /// <param name="errorHandlingMode">The error handling mode.</param>
    /// <param name="token"></param>
    /// <returns></returns>
    private Task<ServiceResponseCollection<ServiceResponse>> ApplyConversationOneTimeAction(
        ConversationActionType actionType,
        IEnumerable<KeyValuePair<ConversationId, DateTime?>> idTimePairs,
        FolderId contextFolderId,
        FolderId? destinationFolderId,
        DeleteMode? deleteType,
        bool? isRead,
        RetentionType? retentionPolicyType,
        Guid? retentionPolicyTagId,
        Flag? flag,
        bool? suppressReadReceipts,
        ServiceErrorHandling errorHandlingMode,
        CancellationToken token
    )
    {
        EwsUtilities.Assert(
            actionType == ConversationActionType.Move ||
            actionType == ConversationActionType.Delete ||
            actionType == ConversationActionType.SetReadState ||
            actionType == ConversationActionType.SetRetentionPolicy ||
            actionType == ConversationActionType.Copy ||
            actionType == ConversationActionType.Flag,
            "ApplyConversationOneTimeAction",
            "Invalid actionType"
        );

        EwsUtilities.ValidateParamCollection(idTimePairs);
        EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2010_SP1, "ApplyConversationAction");

        var request = new ApplyConversationActionRequest(this, errorHandlingMode);

        foreach (var idTimePair in idTimePairs)
        {
            var action = new ConversationAction
            {
                Action = actionType,
                ConversationId = idTimePair.Key,
                ContextFolderId = contextFolderId != null ? new FolderIdWrapper(contextFolderId) : null,
                DestinationFolderId = destinationFolderId != null ? new FolderIdWrapper(destinationFolderId) : null,
                ConversationLastSyncTime = idTimePair.Value,
                IsRead = isRead,
                DeleteType = deleteType,
                RetentionPolicyType = retentionPolicyType,
                RetentionPolicyTagId = retentionPolicyTagId,
                Flag = flag,
                SuppressReadReceipts = suppressReadReceipts,
            };

            request.ConversationActions.Add(action);
        }

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Sets up a conversation so that any item received within that conversation is always categorized.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="conversationId">The id of the conversation.</param>
    /// <param name="categories">The categories that should be stamped on items in the conversation.</param>
    /// <param name="processSynchronously">
    ///     Indicates whether the method should return only once enabling this rule and stamping existing items
    ///     in the conversation is completely done. If processSynchronously is false, the method returns immediately.
    /// </param>
    /// <param name="token"></param>
    /// <returns></returns>
    public Task<ServiceResponseCollection<ServiceResponse>> EnableAlwaysCategorizeItemsInConversations(
        IEnumerable<ConversationId> conversationId,
        IEnumerable<string> categories,
        bool processSynchronously,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParamCollection(categories);

        return ApplyConversationAction(
            ConversationActionType.AlwaysCategorize,
            conversationId,
            processSynchronously,
            new StringList(categories),
            false,
            null,
            ServiceErrorHandling.ReturnErrors,
            token
        );
    }

    /// <summary>
    ///     Sets up a conversation so that any item received within that conversation is no longer categorized.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="conversationId">The id of the conversation.</param>
    /// <param name="processSynchronously">
    ///     Indicates whether the method should return only once disabling this rule and removing the categories from existing
    ///     items
    ///     in the conversation is completely done. If processSynchronously is false, the method returns immediately.
    /// </param>
    /// <param name="token"></param>
    /// <returns></returns>
    public Task<ServiceResponseCollection<ServiceResponse>> DisableAlwaysCategorizeItemsInConversations(
        IEnumerable<ConversationId> conversationId,
        bool processSynchronously,
        CancellationToken token = default
    )
    {
        return ApplyConversationAction(
            ConversationActionType.AlwaysCategorize,
            conversationId,
            processSynchronously,
            null,
            false,
            null,
            ServiceErrorHandling.ReturnErrors,
            token
        );
    }

    /// <summary>
    ///     Sets up a conversation so that any item received within that conversation is always moved to Deleted Items folder.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="conversationId">The id of the conversation.</param>
    /// <param name="processSynchronously">
    ///     Indicates whether the method should return only once enabling this rule and deleting existing items
    ///     in the conversation is completely done. If processSynchronously is false, the method returns immediately.
    /// </param>
    /// <param name="token"></param>
    /// <returns></returns>
    public Task<ServiceResponseCollection<ServiceResponse>> EnableAlwaysDeleteItemsInConversations(
        IEnumerable<ConversationId> conversationId,
        bool processSynchronously,
        CancellationToken token = default
    )
    {
        return ApplyConversationAction(
            ConversationActionType.AlwaysDelete,
            conversationId,
            processSynchronously,
            null,
            true,
            null,
            ServiceErrorHandling.ReturnErrors,
            token
        );
    }

    /// <summary>
    ///     Sets up a conversation so that any item received within that conversation is no longer moved to Deleted Items
    ///     folder.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="conversationId">The id of the conversation.</param>
    /// <param name="processSynchronously">
    ///     Indicates whether the method should return only once disabling this rule and restoring the items
    ///     in the conversation is completely done. If processSynchronously is false, the method returns immediately.
    /// </param>
    /// <param name="token"></param>
    /// <returns></returns>
    public Task<ServiceResponseCollection<ServiceResponse>> DisableAlwaysDeleteItemsInConversations(
        IEnumerable<ConversationId> conversationId,
        bool processSynchronously,
        CancellationToken token = default
    )
    {
        return ApplyConversationAction(
            ConversationActionType.AlwaysDelete,
            conversationId,
            processSynchronously,
            null,
            false,
            null,
            ServiceErrorHandling.ReturnErrors,
            token
        );
    }

    /// <summary>
    ///     Sets up a conversation so that any item received within that conversation is always moved to a specific folder.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="conversationId">The id of the conversation.</param>
    /// <param name="destinationFolderId">The Id of the folder to which conversation items should be moved.</param>
    /// <param name="processSynchronously">
    ///     Indicates whether the method should return only once enabling this rule and moving existing items
    ///     in the conversation is completely done. If processSynchronously is false, the method returns immediately.
    /// </param>
    /// <param name="token"></param>
    /// <returns></returns>
    public Task<ServiceResponseCollection<ServiceResponse>> EnableAlwaysMoveItemsInConversations(
        IEnumerable<ConversationId> conversationId,
        FolderId destinationFolderId,
        bool processSynchronously,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(destinationFolderId);

        return ApplyConversationAction(
            ConversationActionType.AlwaysMove,
            conversationId,
            processSynchronously,
            null,
            false,
            destinationFolderId,
            ServiceErrorHandling.ReturnErrors,
            token
        );
    }

    /// <summary>
    ///     Sets up a conversation so that any item received within that conversation is no longer moved to a specific folder.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="conversationIds">The conversation ids.</param>
    /// <param name="processSynchronously">
    ///     Indicates whether the method should return only once disabling this rule is completely done.
    ///     If processSynchronously is false, the method returns immediately.
    /// </param>
    /// <param name="token"></param>
    /// <returns></returns>
    public Task<ServiceResponseCollection<ServiceResponse>> DisableAlwaysMoveItemsInConversations(
        IEnumerable<ConversationId> conversationIds,
        bool processSynchronously,
        CancellationToken token = default
    )
    {
        return ApplyConversationAction(
            ConversationActionType.AlwaysMove,
            conversationIds,
            processSynchronously,
            null,
            false,
            null,
            ServiceErrorHandling.ReturnErrors,
            token
        );
    }

    /// <summary>
    ///     Moves the items in the specified conversation to the specified destination folder.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="idLastSyncTimePairs">
    ///     The pairs of Id of conversation whose
    ///     items should be moved and the dateTime conversation was last synced
    ///     (Items received after that dateTime will not be moved).
    /// </param>
    /// <param name="contextFolderId">The Id of the folder that contains the conversation.</param>
    /// <param name="destinationFolderId">The Id of the destination folder.</param>
    /// <param name="token"></param>
    /// <returns></returns>
    public Task<ServiceResponseCollection<ServiceResponse>> MoveItemsInConversations(
        IEnumerable<KeyValuePair<ConversationId, DateTime?>> idLastSyncTimePairs,
        FolderId contextFolderId,
        FolderId destinationFolderId,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(destinationFolderId);

        return ApplyConversationOneTimeAction(
            ConversationActionType.Move,
            idLastSyncTimePairs,
            contextFolderId,
            destinationFolderId,
            null,
            null,
            null,
            null,
            null,
            null,
            ServiceErrorHandling.ReturnErrors,
            token
        );
    }

    /// <summary>
    ///     Copies the items in the specified conversation to the specified destination folder.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="idLastSyncTimePairs">
    ///     The pairs of Id of conversation whose
    ///     items should be copied and the date and time conversation was last synced
    ///     (Items received after that date will not be copied).
    /// </param>
    /// <param name="contextFolderId">The context folder id.</param>
    /// <param name="destinationFolderId">The destination folder id.</param>
    /// <param name="token"></param>
    /// <returns></returns>
    public Task<ServiceResponseCollection<ServiceResponse>> CopyItemsInConversations(
        IEnumerable<KeyValuePair<ConversationId, DateTime?>> idLastSyncTimePairs,
        FolderId contextFolderId,
        FolderId destinationFolderId,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(destinationFolderId);

        return ApplyConversationOneTimeAction(
            ConversationActionType.Copy,
            idLastSyncTimePairs,
            contextFolderId,
            destinationFolderId,
            null,
            null,
            null,
            null,
            null,
            null,
            ServiceErrorHandling.ReturnErrors,
            token
        );
    }

    /// <summary>
    ///     Deletes the items in the specified conversation. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="idLastSyncTimePairs">
    ///     The pairs of Id of conversation whose
    ///     items should be deleted and the date and time conversation was last synced
    ///     (Items received after that date will not be deleted).
    /// </param>
    /// <param name="contextFolderId">The Id of the folder that contains the conversation.</param>
    /// <param name="deleteMode">The deletion mode.</param>
    /// <param name="token"></param>
    /// <returns></returns>
    public Task<ServiceResponseCollection<ServiceResponse>> DeleteItemsInConversations(
        IEnumerable<KeyValuePair<ConversationId, DateTime?>> idLastSyncTimePairs,
        FolderId contextFolderId,
        DeleteMode deleteMode,
        CancellationToken token = default
    )
    {
        return ApplyConversationOneTimeAction(
            ConversationActionType.Delete,
            idLastSyncTimePairs,
            contextFolderId,
            null,
            deleteMode,
            null,
            null,
            null,
            null,
            null,
            ServiceErrorHandling.ReturnErrors,
            token
        );
    }

    /// <summary>
    ///     Sets the read state for items in conversation. Calling this method would
    ///     result in call to EWS.
    /// </summary>
    /// <param name="idLastSyncTimePairs">
    ///     The pairs of Id of conversation whose
    ///     items should have their read state set and the date and time conversation
    ///     was last synced (Items received after that date will not have their read
    ///     state set).
    /// </param>
    /// <param name="contextFolderId">The Id of the folder that contains the conversation.</param>
    /// <param name="isRead">if set to <c>true</c>, conversation items are marked as read; otherwise they are marked as unread.</param>
    /// <param name="token"></param>
    public Task<ServiceResponseCollection<ServiceResponse>> SetReadStateForItemsInConversations(
        IEnumerable<KeyValuePair<ConversationId, DateTime?>> idLastSyncTimePairs,
        FolderId contextFolderId,
        bool isRead,
        CancellationToken token = default
    )
    {
        return ApplyConversationOneTimeAction(
            ConversationActionType.SetReadState,
            idLastSyncTimePairs,
            contextFolderId,
            null,
            null,
            isRead,
            null,
            null,
            null,
            null,
            ServiceErrorHandling.ReturnErrors,
            token
        );
    }

    /// <summary>
    ///     Sets the read state for items in conversation. Calling this method would
    ///     result in call to EWS.
    /// </summary>
    /// <param name="idLastSyncTimePairs">
    ///     The pairs of Id of conversation whose
    ///     items should have their read state set and the date and time conversation
    ///     was last synced (Items received after that date will not have their read
    ///     state set).
    /// </param>
    /// <param name="contextFolderId">The Id of the folder that contains the conversation.</param>
    /// <param name="isRead">if set to <c>true</c>, conversation items are marked as read; otherwise they are marked as unread.</param>
    /// <param name="suppressReadReceipts">if set to <c>true</c> read receipts are suppressed.</param>
    /// <param name="token"></param>
    public Task<ServiceResponseCollection<ServiceResponse>> SetReadStateForItemsInConversations(
        IEnumerable<KeyValuePair<ConversationId, DateTime?>> idLastSyncTimePairs,
        FolderId contextFolderId,
        bool isRead,
        bool suppressReadReceipts,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013, "SetReadStateForItemsInConversations");

        return ApplyConversationOneTimeAction(
            ConversationActionType.SetReadState,
            idLastSyncTimePairs,
            contextFolderId,
            null,
            null,
            isRead,
            null,
            null,
            null,
            suppressReadReceipts,
            ServiceErrorHandling.ReturnErrors,
            token
        );
    }

    /// <summary>
    ///     Sets the retention policy for items in conversation. Calling this method would
    ///     result in call to EWS.
    /// </summary>
    /// <param name="idLastSyncTimePairs">
    ///     The pairs of Id of conversation whose
    ///     items should have their retention policy set and the date and time conversation
    ///     was last synced (Items received after that date will not have their retention
    ///     policy set).
    /// </param>
    /// <param name="contextFolderId">The Id of the folder that contains the conversation.</param>
    /// <param name="retentionPolicyType">Retention policy type.</param>
    /// <param name="retentionPolicyTagId">Retention policy tag id.  Null will clear the policy.</param>
    /// <param name="token"></param>
    public Task<ServiceResponseCollection<ServiceResponse>> SetRetentionPolicyForItemsInConversations(
        IEnumerable<KeyValuePair<ConversationId, DateTime?>> idLastSyncTimePairs,
        FolderId contextFolderId,
        RetentionType retentionPolicyType,
        Guid? retentionPolicyTagId,
        CancellationToken token = default
    )
    {
        return ApplyConversationOneTimeAction(
            ConversationActionType.SetRetentionPolicy,
            idLastSyncTimePairs,
            contextFolderId,
            null,
            null,
            null,
            retentionPolicyType,
            retentionPolicyTagId,
            null,
            null,
            ServiceErrorHandling.ReturnErrors,
            token
        );
    }

    /// <summary>
    ///     Sets flag status for items in conversation. Calling this method would result in call to EWS.
    /// </summary>
    /// <param name="idLastSyncTimePairs">
    ///     The pairs of Id of conversation whose
    ///     items should have their read state set and the date and time conversation
    ///     was last synced (Items received after that date will not have their read
    ///     state set).
    /// </param>
    /// <param name="contextFolderId">The Id of the folder that contains the conversation.</param>
    /// <param name="flagStatus">Flag status to apply to conversation items.</param>
    /// <param name="token"></param>
    public Task<ServiceResponseCollection<ServiceResponse>> SetFlagStatusForItemsInConversations(
        IEnumerable<KeyValuePair<ConversationId, DateTime?>> idLastSyncTimePairs,
        FolderId contextFolderId,
        Flag flagStatus,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013, "SetFlagStatusForItemsInConversations");

        return ApplyConversationOneTimeAction(
            ConversationActionType.Flag,
            idLastSyncTimePairs,
            contextFolderId,
            null,
            null,
            null,
            null,
            null,
            flagStatus,
            null,
            ServiceErrorHandling.ReturnErrors,
            token
        );
    }

    #endregion


    #region Id conversion operations

    /// <summary>
    ///     Converts multiple Ids from one format to another in a single call to EWS.
    /// </summary>
    /// <param name="ids">The Ids to convert.</param>
    /// <param name="destinationFormat">The destination format.</param>
    /// <param name="errorHandling">Type of error handling to perform.</param>
    /// <param name="token"></param>
    /// <returns>A ServiceResponseCollection providing conversion results for each specified Ids.</returns>
    private Task<ServiceResponseCollection<ConvertIdResponse>> InternalConvertIds(
        IEnumerable<AlternateIdBase> ids,
        IdFormat destinationFormat,
        ServiceErrorHandling errorHandling,
        CancellationToken token
    )
    {
        EwsUtilities.ValidateParamCollection(ids);

        var request = new ConvertIdRequest(this, errorHandling)
        {
            DestinationFormat = destinationFormat,
        };

        request.Ids.AddRange(ids);

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Converts multiple Ids from one format to another in a single call to EWS.
    /// </summary>
    /// <param name="ids">The Ids to convert.</param>
    /// <param name="destinationFormat">The destination format.</param>
    /// <param name="token"></param>
    /// <returns>A ServiceResponseCollection providing conversion results for each specified Ids.</returns>
    public Task<ServiceResponseCollection<ConvertIdResponse>> ConvertIds(
        IEnumerable<AlternateIdBase> ids,
        IdFormat destinationFormat,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParamCollection(ids);

        return InternalConvertIds(ids, destinationFormat, ServiceErrorHandling.ReturnErrors, token);
    }

    /// <summary>
    ///     Converts Id from one format to another in a single call to EWS.
    /// </summary>
    /// <param name="id">The Id to convert.</param>
    /// <param name="destinationFormat">The destination format.</param>
    /// <param name="token"></param>
    /// <returns>The converted Id.</returns>
    public async Task<AlternateIdBase> ConvertId(
        AlternateIdBase id,
        IdFormat destinationFormat,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(id);

        var responses = await InternalConvertIds([id,], destinationFormat, ServiceErrorHandling.ThrowOnError, token)
            .ConfigureAwait(false);

        return responses[0].ConvertedId;
    }

    #endregion


    #region Delegate management operations

    /// <summary>
    ///     Adds delegates to a specific mailbox. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="mailbox">The mailbox to add delegates to.</param>
    /// <param name="meetingRequestsDeliveryScope">Indicates how meeting requests should be sent to delegates.</param>
    /// <param name="delegateUsers">The delegate users to add.</param>
    /// <returns>A collection of DelegateUserResponse objects providing the results of the operation.</returns>
    public Task<Collection<DelegateUserResponse>> AddDelegates(
        Mailbox mailbox,
        MeetingRequestsDeliveryScope? meetingRequestsDeliveryScope,
        params DelegateUser[] delegateUsers
    )
    {
        return AddDelegates(mailbox, meetingRequestsDeliveryScope, (IEnumerable<DelegateUser>)delegateUsers);
    }

    /// <summary>
    ///     Adds delegates to a specific mailbox. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="mailbox">The mailbox to add delegates to.</param>
    /// <param name="meetingRequestsDeliveryScope">Indicates how meeting requests should be sent to delegates.</param>
    /// <param name="delegateUsers">The delegate users to add.</param>
    /// <param name="token"></param>
    /// <returns>A collection of DelegateUserResponse objects providing the results of the operation.</returns>
    public async Task<Collection<DelegateUserResponse>> AddDelegates(
        Mailbox mailbox,
        MeetingRequestsDeliveryScope? meetingRequestsDeliveryScope,
        IEnumerable<DelegateUser> delegateUsers,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(mailbox);
        EwsUtilities.ValidateParamCollection(delegateUsers);

        var request = new AddDelegateRequest(this)
        {
            Mailbox = mailbox,
            MeetingRequestsDeliveryScope = meetingRequestsDeliveryScope,
        };

        request.DelegateUsers.AddRange(delegateUsers);

        var response = await request.Execute(token).ConfigureAwait(false);
        return response.DelegateUserResponses;
    }

    /// <summary>
    ///     Updates delegates on a specific mailbox. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="mailbox">The mailbox to update delegates on.</param>
    /// <param name="meetingRequestsDeliveryScope">Indicates how meeting requests should be sent to delegates.</param>
    /// <param name="token"></param>
    /// <param name="delegateUsers">The delegate users to update.</param>
    /// <returns>A collection of DelegateUserResponse objects providing the results of the operation.</returns>
    public Task<Collection<DelegateUserResponse>> UpdateDelegates(
        Mailbox mailbox,
        MeetingRequestsDeliveryScope? meetingRequestsDeliveryScope,
        CancellationToken token = default,
        params DelegateUser[] delegateUsers
    )
    {
        return UpdateDelegates(mailbox, meetingRequestsDeliveryScope, delegateUsers, token);
    }

    /// <summary>
    ///     Updates delegates on a specific mailbox. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="mailbox">The mailbox to update delegates on.</param>
    /// <param name="meetingRequestsDeliveryScope">Indicates how meeting requests should be sent to delegates.</param>
    /// <param name="delegateUsers">The delegate users to update.</param>
    /// <param name="token"></param>
    /// <returns>A collection of DelegateUserResponse objects providing the results of the operation.</returns>
    public async Task<Collection<DelegateUserResponse>> UpdateDelegates(
        Mailbox mailbox,
        MeetingRequestsDeliveryScope? meetingRequestsDeliveryScope,
        IEnumerable<DelegateUser> delegateUsers,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(mailbox);
        EwsUtilities.ValidateParamCollection(delegateUsers);

        var request = new UpdateDelegateRequest(this)
        {
            Mailbox = mailbox,
            MeetingRequestsDeliveryScope = meetingRequestsDeliveryScope,
        };

        request.DelegateUsers.AddRange(delegateUsers);

        var response = await request.Execute(token).ConfigureAwait(false);
        return response.DelegateUserResponses;
    }

    /// <summary>
    ///     Removes delegates on a specific mailbox. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="mailbox">The mailbox to remove delegates from.</param>
    /// <param name="token"></param>
    /// <param name="userIds">The Ids of the delegate users to remove.</param>
    /// <returns>A collection of DelegateUserResponse objects providing the results of the operation.</returns>
    public Task<Collection<DelegateUserResponse>> RemoveDelegates(
        Mailbox mailbox,
        CancellationToken token = default,
        params UserId[] userIds
    )
    {
        return RemoveDelegates(mailbox, userIds, token);
    }

    /// <summary>
    ///     Removes delegates on a specific mailbox. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="mailbox">The mailbox to remove delegates from.</param>
    /// <param name="userIds">The Ids of the delegate users to remove.</param>
    /// <param name="token"></param>
    /// <returns>A collection of DelegateUserResponse objects providing the results of the operation.</returns>
    public async Task<Collection<DelegateUserResponse>> RemoveDelegates(
        Mailbox mailbox,
        IEnumerable<UserId> userIds,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(mailbox);
        EwsUtilities.ValidateParamCollection(userIds);

        var request = new RemoveDelegateRequest(this)
        {
            Mailbox = mailbox,
        };

        request.UserIds.AddRange(userIds);

        var response = await request.Execute(token).ConfigureAwait(false);
        return response.DelegateUserResponses;
    }

    /// <summary>
    ///     Retrieves the delegates of a specific mailbox. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="mailbox">The mailbox to retrieve the delegates of.</param>
    /// <param name="includePermissions">Indicates whether detailed permissions should be returned fro each delegate.</param>
    /// <param name="token"></param>
    /// <param name="userIds">The optional Ids of the delegate users to retrieve.</param>
    /// <returns>A GetDelegateResponse providing the results of the operation.</returns>
    public Task<DelegateInformation> GetDelegates(
        Mailbox mailbox,
        bool includePermissions,
        CancellationToken token = default,
        params UserId[] userIds
    )
    {
        return GetDelegates(mailbox, includePermissions, userIds, token);
    }

    /// <summary>
    ///     Retrieves the delegates of a specific mailbox. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="mailbox">The mailbox to retrieve the delegates of.</param>
    /// <param name="includePermissions">Indicates whether detailed permissions should be returned fro each delegate.</param>
    /// <param name="userIds">The optional Ids of the delegate users to retrieve.</param>
    /// <param name="token"></param>
    /// <returns>A GetDelegateResponse providing the results of the operation.</returns>
    public async Task<DelegateInformation> GetDelegates(
        Mailbox mailbox,
        bool includePermissions,
        IEnumerable<UserId> userIds,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(mailbox);

        var request = new GetDelegateRequest(this)
        {
            Mailbox = mailbox,
            IncludePermissions = includePermissions,
        };

        request.UserIds.AddRange(userIds);

        var response = await request.Execute(token).ConfigureAwait(false);
        var delegateInformation = new DelegateInformation(
            response.DelegateUserResponses,
            response.MeetingRequestsDeliveryScope
        );

        return delegateInformation;
    }

    #endregion


    #region UserConfiguration operations

    /// <summary>
    ///     Creates a UserConfiguration.
    /// </summary>
    /// <param name="userConfiguration">The UserConfiguration.</param>
    /// <param name="token"></param>
    internal System.Threading.Tasks.Task CreateUserConfiguration(
        UserConfiguration userConfiguration,
        CancellationToken token
    )
    {
        EwsUtilities.ValidateParam(userConfiguration);

        var request = new CreateUserConfigurationRequest(this)
        {
            UserConfiguration = userConfiguration,
        };

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Deletes a UserConfiguration.
    /// </summary>
    /// <param name="name">Name of the UserConfiguration to retrieve.</param>
    /// <param name="parentFolderId">Id of the folder containing the UserConfiguration.</param>
    /// <param name="token"></param>
    internal System.Threading.Tasks.Task DeleteUserConfiguration(
        string name,
        FolderId parentFolderId,
        CancellationToken token
    )
    {
        EwsUtilities.ValidateParam(name);
        EwsUtilities.ValidateParam(parentFolderId);

        var request = new DeleteUserConfigurationRequest(this)
        {
            Name = name,
            ParentFolderId = parentFolderId,
        };

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Gets a UserConfiguration.
    /// </summary>
    /// <param name="name">Name of the UserConfiguration to retrieve.</param>
    /// <param name="parentFolderId">Id of the folder containing the UserConfiguration.</param>
    /// <param name="properties">Properties to retrieve.</param>
    /// <param name="token"></param>
    /// <returns>A UserConfiguration.</returns>
    internal async Task<UserConfiguration> GetUserConfiguration(
        string name,
        FolderId parentFolderId,
        UserConfigurationProperties properties,
        CancellationToken token
    )
    {
        EwsUtilities.ValidateParam(name);
        EwsUtilities.ValidateParam(parentFolderId);

        var request = new GetUserConfigurationRequest(this)
        {
            Name = name,
            ParentFolderId = parentFolderId,
            Properties = properties,
        };

        var responses = await request.ExecuteAsync(token).ConfigureAwait(false);
        return responses[0].UserConfiguration;
    }

    /// <summary>
    ///     Loads the properties of the specified userConfiguration.
    /// </summary>
    /// <param name="userConfiguration">The userConfiguration containing properties to load.</param>
    /// <param name="properties">Properties to retrieve.</param>
    /// <param name="token"></param>
    internal System.Threading.Tasks.Task LoadPropertiesForUserConfiguration(
        UserConfiguration userConfiguration,
        UserConfigurationProperties properties,
        CancellationToken token
    )
    {
        EwsUtilities.Assert(
            userConfiguration != null,
            "ExchangeService.LoadPropertiesForUserConfiguration",
            "userConfiguration is null"
        );

        var request = new GetUserConfigurationRequest(this)
        {
            UserConfiguration = userConfiguration,
            Properties = properties,
        };

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Updates a UserConfiguration.
    /// </summary>
    /// <param name="userConfiguration">The UserConfiguration.</param>
    /// <param name="token"></param>
    internal System.Threading.Tasks.Task UpdateUserConfiguration(
        UserConfiguration userConfiguration,
        CancellationToken token
    )
    {
        EwsUtilities.ValidateParam(userConfiguration);

        var request = new UpdateUserConfigurationRequest(this)
        {
            UserConfiguration = userConfiguration,
        };

        return request.ExecuteAsync(token);
    }

    #endregion


    #region InboxRule operations

    /// <summary>
    ///     Retrieves inbox rules of the authenticated user.
    /// </summary>
    /// <returns>A RuleCollection object containing the authenticated user's inbox rules.</returns>
    public async Task<RuleCollection> GetInboxRules(CancellationToken token = default)
    {
        var request = new GetInboxRulesRequest(this);

        var response = await request.Execute(token).ConfigureAwait(false);
        return response.Rules;
    }

    /// <summary>
    ///     Retrieves the inbox rules of the specified user.
    /// </summary>
    /// <param name="mailboxSmtpAddress">The SMTP address of the user whose inbox rules should be retrieved.</param>
    /// <param name="token"></param>
    /// <returns>A RuleCollection object containing the inbox rules of the specified user.</returns>
    public async Task<RuleCollection> GetInboxRules(string mailboxSmtpAddress, CancellationToken token = default)
    {
        EwsUtilities.ValidateParam(mailboxSmtpAddress, "MailboxSmtpAddress");

        var request = new GetInboxRulesRequest(this)
        {
            MailboxSmtpAddress = mailboxSmtpAddress,
        };

        var response = await request.Execute(token).ConfigureAwait(false);
        return response.Rules;
    }

    /// <summary>
    ///     Updates the authenticated user's inbox rules by applying the specified operations.
    /// </summary>
    /// <param name="operations">The operations that should be applied to the user's inbox rules.</param>
    /// <param name="removeOutlookRuleBlob">Indicate whether or not to remove Outlook Rule Blob.</param>
    /// <param name="token"></param>
    public System.Threading.Tasks.Task UpdateInboxRules(
        IEnumerable<RuleOperation> operations,
        bool removeOutlookRuleBlob,
        CancellationToken token = default
    )
    {
        var request = new UpdateInboxRulesRequest(this)
        {
            InboxRuleOperations = operations,
            RemoveOutlookRuleBlob = removeOutlookRuleBlob,
        };
        return request.Execute(token);
    }

    /// <summary>
    ///     Update the specified user's inbox rules by applying the specified operations.
    /// </summary>
    /// <param name="operations">The operations that should be applied to the user's inbox rules.</param>
    /// <param name="removeOutlookRuleBlob">Indicate whether or not to remove Outlook Rule Blob.</param>
    /// <param name="mailboxSmtpAddress">The SMTP address of the user whose inbox rules should be updated.</param>
    /// <param name="token"></param>
    public System.Threading.Tasks.Task UpdateInboxRules(
        IEnumerable<RuleOperation> operations,
        bool removeOutlookRuleBlob,
        string mailboxSmtpAddress,
        CancellationToken token = default
    )
    {
        var request = new UpdateInboxRulesRequest(this)
        {
            InboxRuleOperations = operations,
            RemoveOutlookRuleBlob = removeOutlookRuleBlob,
            MailboxSmtpAddress = mailboxSmtpAddress,
        };
        return request.Execute(token);
    }

    #endregion


    #region eDiscovery/Compliance operations

    /// <summary>
    ///     Get discovery search configuration
    /// </summary>
    /// <param name="searchId">Search Id</param>
    /// <param name="expandGroupMembership">True if want to expand group membership</param>
    /// <param name="inPlaceHoldConfigurationOnly">True if only want the inplacehold configuration</param>
    /// <param name="token"></param>
    /// <returns>Service response object</returns>
    public Task<GetDiscoverySearchConfigurationResponse> GetDiscoverySearchConfiguration(
        string searchId,
        bool expandGroupMembership,
        bool inPlaceHoldConfigurationOnly,
        CancellationToken token = default
    )
    {
        var request = new GetDiscoverySearchConfigurationRequest(this)
        {
            SearchId = searchId,
            ExpandGroupMembership = expandGroupMembership,
            InPlaceHoldConfigurationOnly = inPlaceHoldConfigurationOnly,
        };

        return request.Execute(token);
    }

    /// <summary>
    ///     Get searchable mailboxes
    /// </summary>
    /// <param name="searchFilter">Search filter</param>
    /// <param name="expandGroupMembership">True if want to expand group membership</param>
    /// <param name="token"></param>
    /// <returns>Service response object</returns>
    public Task<GetSearchableMailboxesResponse> GetSearchableMailboxes(
        string searchFilter,
        bool expandGroupMembership,
        CancellationToken token = default
    )
    {
        var request = new GetSearchableMailboxesRequest(this)
        {
            SearchFilter = searchFilter,
            ExpandGroupMembership = expandGroupMembership,
        };

        return request.Execute(token);
    }

    /// <summary>
    ///     Search mailboxes
    /// </summary>
    /// <param name="mailboxQueries">Collection of query and mailboxes</param>
    /// <param name="resultType">Search result type</param>
    /// <param name="token"></param>
    /// <returns>Collection of search mailboxes response object</returns>
    public Task<ServiceResponseCollection<SearchMailboxesResponse>> SearchMailboxes(
        IEnumerable<MailboxQuery>? mailboxQueries,
        SearchResultType resultType,
        CancellationToken token = default
    )
    {
        var request = new SearchMailboxesRequest(this, ServiceErrorHandling.ReturnErrors)
        {
            ResultType = resultType,
        };

        if (mailboxQueries != null)
        {
            request.SearchQueries.AddRange(mailboxQueries);
        }

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Search mailboxes
    /// </summary>
    /// <param name="mailboxQueries">Collection of query and mailboxes</param>
    /// <param name="resultType">Search result type</param>
    /// <param name="sortByProperty">Sort by property name</param>
    /// <param name="sortOrder">Sort order</param>
    /// <param name="pageSize">Page size</param>
    /// <param name="pageDirection">Page navigation direction</param>
    /// <param name="pageItemReference">Item reference used for paging</param>
    /// <param name="token"></param>
    /// <returns>Collection of search mailboxes response object</returns>
    public Task<ServiceResponseCollection<SearchMailboxesResponse>> SearchMailboxes(
        IEnumerable<MailboxQuery>? mailboxQueries,
        SearchResultType resultType,
        string sortByProperty,
        SortDirection sortOrder,
        int pageSize,
        SearchPageDirection pageDirection,
        string pageItemReference,
        CancellationToken token = default
    )
    {
        var request = new SearchMailboxesRequest(this, ServiceErrorHandling.ReturnErrors)
        {
            ResultType = resultType,
            SortByProperty = sortByProperty,
            SortOrder = sortOrder,
            PageSize = pageSize,
            PageDirection = pageDirection,
            PageItemReference = pageItemReference,
        };

        if (mailboxQueries != null)
        {
            request.SearchQueries.AddRange(mailboxQueries);
        }

        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Search mailboxes
    /// </summary>
    /// <param name="searchParameters">Search mailboxes parameters</param>
    /// <param name="token"></param>
    /// <returns>Collection of search mailboxes response object</returns>
    public Task<ServiceResponseCollection<SearchMailboxesResponse>> SearchMailboxes(
        SearchMailboxesParameters searchParameters,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(searchParameters);
        EwsUtilities.ValidateParam(searchParameters.SearchQueries);

        var request = CreateSearchMailboxesRequest(searchParameters);
        return request.ExecuteAsync(token);
    }

    /// <summary>
    ///     Set hold on mailboxes
    /// </summary>
    /// <param name="holdId">Hold id</param>
    /// <param name="actionType">Action type</param>
    /// <param name="query">Query string</param>
    /// <param name="mailboxes">Collection of mailboxes</param>
    /// <param name="token"></param>
    /// <returns>Service response object</returns>
    public Task<SetHoldOnMailboxesResponse> SetHoldOnMailboxes(
        string holdId,
        HoldAction actionType,
        string query,
        string[] mailboxes,
        CancellationToken token = default
    )
    {
        var request = new SetHoldOnMailboxesRequest(this)
        {
            HoldId = holdId,
            ActionType = actionType,
            Query = query,
            Mailboxes = mailboxes,
            InPlaceHoldIdentity = null,
        };

        return request.Execute(token);
    }

    /// <summary>
    ///     Set hold on mailboxes
    /// </summary>
    /// <param name="holdId">Hold id</param>
    /// <param name="actionType">Action type</param>
    /// <param name="query">Query string</param>
    /// <param name="inPlaceHoldIdentity">in-place hold identity</param>
    /// <returns>Service response object</returns>
    public Task<SetHoldOnMailboxesResponse> SetHoldOnMailboxes(
        string holdId,
        HoldAction actionType,
        string query,
        string inPlaceHoldIdentity
    )
    {
        return SetHoldOnMailboxes(holdId, actionType, query, inPlaceHoldIdentity, null);
    }

    /// <summary>
    ///     Set hold on mailboxes
    /// </summary>
    /// <param name="holdId">Hold id</param>
    /// <param name="actionType">Action type</param>
    /// <param name="query">Query string</param>
    /// <param name="inPlaceHoldIdentity">in-place hold identity</param>
    /// <param name="itemHoldPeriod">item hold period</param>
    /// <param name="token"></param>
    /// <returns>Service response object</returns>
    public Task<SetHoldOnMailboxesResponse> SetHoldOnMailboxes(
        string holdId,
        HoldAction actionType,
        string query,
        string inPlaceHoldIdentity,
        string? itemHoldPeriod,
        CancellationToken token = default
    )
    {
        var request = new SetHoldOnMailboxesRequest(this)
        {
            HoldId = holdId,
            ActionType = actionType,
            Query = query,
            Mailboxes = null,
            InPlaceHoldIdentity = inPlaceHoldIdentity,
            ItemHoldPeriod = itemHoldPeriod,
        };

        return request.Execute(token);
    }

    /// <summary>
    ///     Set hold on mailboxes
    /// </summary>
    /// <param name="parameters">Set hold parameters</param>
    /// <param name="token"></param>
    /// <returns>Service response object</returns>
    public Task<SetHoldOnMailboxesResponse> SetHoldOnMailboxes(
        SetHoldOnMailboxesParameters parameters,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(parameters);

        var request = new SetHoldOnMailboxesRequest(this)
        {
            HoldId = parameters.HoldId,
            ActionType = parameters.ActionType,
            Query = parameters.Query,
            Mailboxes = parameters.Mailboxes,
            Language = parameters.Language,
            InPlaceHoldIdentity = parameters.InPlaceHoldIdentity,
        };

        return request.Execute(token);
    }

    /// <summary>
    ///     Get hold on mailboxes
    /// </summary>
    /// <param name="holdId">Hold id</param>
    /// <param name="token"></param>
    /// <returns>Service response object</returns>
    public Task<GetHoldOnMailboxesResponse> GetHoldOnMailboxes(string holdId, CancellationToken token = default)
    {
        var request = new GetHoldOnMailboxesRequest(this)
        {
            HoldId = holdId,
        };

        return request.Execute(token);
    }

    /// <summary>
    ///     Get non indexable item details
    /// </summary>
    /// <param name="mailboxes">Array of mailbox legacy DN</param>
    /// <returns>Service response object</returns>
    public Task<GetNonIndexableItemDetailsResponse> GetNonIndexableItemDetails(string[] mailboxes)
    {
        return GetNonIndexableItemDetails(mailboxes, pageSize: null, pageItemReference: null, pageDirection: null);
    }

    /// <summary>
    ///     Get non indexable item details
    /// </summary>
    /// <param name="mailboxes">Array of mailbox legacy DN</param>
    /// <param name="pageSize">The page size</param>
    /// <param name="pageItemReference">Page item reference</param>
    /// <param name="pageDirection">Page direction</param>
    /// <returns>Service response object</returns>
    public Task<GetNonIndexableItemDetailsResponse> GetNonIndexableItemDetails(
        string[] mailboxes,
        int? pageSize,
        string? pageItemReference,
        SearchPageDirection? pageDirection
    )
    {
        var parameters = new GetNonIndexableItemDetailsParameters
        {
            Mailboxes = mailboxes,
            PageSize = pageSize,
            PageItemReference = pageItemReference,
            PageDirection = pageDirection,
            SearchArchiveOnly = false,
        };

        return GetNonIndexableItemDetails(parameters);
    }

    /// <summary>
    ///     Get non indexable item details
    /// </summary>
    /// <param name="parameters">Get non indexable item details parameters</param>
    /// <param name="token"></param>
    /// <returns>Service response object</returns>
    public Task<GetNonIndexableItemDetailsResponse> GetNonIndexableItemDetails(
        GetNonIndexableItemDetailsParameters parameters,
        CancellationToken token = default
    )
    {
        var request = CreateGetNonIndexableItemDetailsRequest(parameters);

        return request.Execute(token);
    }

    /// <summary>
    ///     Get non indexable item statistics
    /// </summary>
    /// <param name="mailboxes">Array of mailbox legacy DN</param>
    /// <returns>Service response object</returns>
    public Task<GetNonIndexableItemStatisticsResponse> GetNonIndexableItemStatistics(string[] mailboxes)
    {
        var parameters = new GetNonIndexableItemStatisticsParameters
        {
            Mailboxes = mailboxes,
            SearchArchiveOnly = false,
        };

        return GetNonIndexableItemStatistics(parameters);
    }

    /// <summary>
    ///     Get non indexable item statistics
    /// </summary>
    /// <param name="parameters">Get non indexable item statistics parameters</param>
    /// <param name="token"></param>
    /// <returns>Service response object</returns>
    public Task<GetNonIndexableItemStatisticsResponse> GetNonIndexableItemStatistics(
        GetNonIndexableItemStatisticsParameters parameters,
        CancellationToken token = default
    )
    {
        var request = CreateGetNonIndexableItemStatisticsRequest(parameters);

        return request.Execute(token);
    }

    /// <summary>
    ///     Create get non indexable item details request
    /// </summary>
    /// <param name="parameters">Get non indexable item details parameters</param>
    /// <returns>GetNonIndexableItemDetails request</returns>
    private GetNonIndexableItemDetailsRequest CreateGetNonIndexableItemDetailsRequest(
        GetNonIndexableItemDetailsParameters parameters
    )
    {
        EwsUtilities.ValidateParam(parameters);
        EwsUtilities.ValidateParam(parameters.Mailboxes);

        var request = new GetNonIndexableItemDetailsRequest(this)
        {
            Mailboxes = parameters.Mailboxes,
            PageSize = parameters.PageSize,
            PageItemReference = parameters.PageItemReference,
            PageDirection = parameters.PageDirection,
            SearchArchiveOnly = parameters.SearchArchiveOnly,
        };

        return request;
    }

    /// <summary>
    ///     Create get non indexable item statistics request
    /// </summary>
    /// <param name="parameters">Get non indexable item statistics parameters</param>
    /// <returns>Service response object</returns>
    private GetNonIndexableItemStatisticsRequest CreateGetNonIndexableItemStatisticsRequest(
        GetNonIndexableItemStatisticsParameters parameters
    )
    {
        EwsUtilities.ValidateParam(parameters);
        EwsUtilities.ValidateParam(parameters.Mailboxes);

        var request = new GetNonIndexableItemStatisticsRequest(this)
        {
            Mailboxes = parameters.Mailboxes,
            SearchArchiveOnly = parameters.SearchArchiveOnly,
        };

        return request;
    }

    /// <summary>
    ///     Creates SearchMailboxesRequest from SearchMailboxesParameters
    /// </summary>
    /// <param name="searchParameters">search parameters</param>
    /// <returns>request object</returns>
    private SearchMailboxesRequest CreateSearchMailboxesRequest(SearchMailboxesParameters searchParameters)
    {
        var request = new SearchMailboxesRequest(this, ServiceErrorHandling.ReturnErrors)
        {
            ResultType = searchParameters.ResultType,
            PreviewItemResponseShape = searchParameters.PreviewItemResponseShape,
            SortByProperty = searchParameters.SortBy,
            SortOrder = searchParameters.SortOrder,
            Language = searchParameters.Language,
            PerformDeduplication = searchParameters.PerformDeduplication,
            PageSize = searchParameters.PageSize,
            PageDirection = searchParameters.PageDirection,
            PageItemReference = searchParameters.PageItemReference,
        };

        request.SearchQueries.AddRange(searchParameters.SearchQueries);
        return request;
    }

    #endregion


    #region ClientAccessTokens

    /// <summary>
    ///     GetClientAccessToken
    /// </summary>
    /// <param name="idAndTypes">Id and Types</param>
    /// <returns>A ServiceResponseCollection providing token results for each of the specified id and types.</returns>
    public Task<ServiceResponseCollection<GetClientAccessTokenResponse>> GetClientAccessToken(
        IEnumerable<KeyValuePair<string, ClientAccessTokenType>> idAndTypes
    )
    {
        // TODO: check this mutation
        _ = new GetClientAccessTokenRequest(this, ServiceErrorHandling.ReturnErrors);

        var requestList = new List<ClientAccessTokenRequest>();
        foreach (var idAndType in idAndTypes)
        {
            var clientAccessTokenRequest = new ClientAccessTokenRequest(idAndType.Key, idAndType.Value);
            requestList.Add(clientAccessTokenRequest);
        }

        return GetClientAccessToken(requestList.ToArray());
    }

    /// <summary>
    ///     GetClientAccessToken
    /// </summary>
    /// <param name="tokenRequests">Token requests array</param>
    /// <param name="token"></param>
    /// <returns>A ServiceResponseCollection providing token results for each of the specified id and types.</returns>
    public Task<ServiceResponseCollection<GetClientAccessTokenResponse>> GetClientAccessToken(
        ClientAccessTokenRequest[] tokenRequests,
        CancellationToken token = default
    )
    {
        var request = new GetClientAccessTokenRequest(this, ServiceErrorHandling.ReturnErrors)
        {
            TokenRequests = tokenRequests,
        };

        return request.ExecuteAsync(token);
    }

    #endregion


    #region Client Extensibility

    /// <summary>
    ///     Get the app manifests.
    /// </summary>
    /// <returns>Collection of manifests</returns>
    public async Task<Collection<XmlDocument>> GetAppManifests(CancellationToken token = default)
    {
        var request = new GetAppManifestsRequest(this);

        var response = await request.Execute(token).ConfigureAwait(false);
        return response.Manifests;
    }

    /// <summary>
    ///     Get the app manifests.  Works with Exchange 2013 SP1 or later EWS.
    /// </summary>
    /// <param name="apiVersionSupported">The api version supported by the client.</param>
    /// <param name="schemaVersionSupported">The schema version supported by the client.</param>
    /// <param name="token"></param>
    /// <returns>Collection of manifests</returns>
    public async Task<Collection<ClientApp>> GetAppManifests(
        string apiVersionSupported,
        string schemaVersionSupported,
        CancellationToken token = default
    )
    {
        var request = new GetAppManifestsRequest(this)
        {
            ApiVersionSupported = apiVersionSupported,
            SchemaVersionSupported = schemaVersionSupported,
        };

        var response = await request.Execute(token).ConfigureAwait(false);
        return response.Apps;
    }

    /// <summary>
    ///     Install App.
    /// </summary>
    /// <param name="manifestStream">
    ///     The manifest's plain text XML stream.
    ///     Notice: Stream has state. If you want this function read from the expected position of the stream,
    ///     please make sure set read position by manifestStream.Position = expectedPosition.
    ///     Be aware read manifestStream.Length puts stream's Position at stream end.
    ///     If you retrieve manifestStream.Length before call this function, nothing will be read.
    ///     When this function succeeds, manifestStream is closed. This is by EWS design to
    ///     release resource in timely manner.
    /// </param>
    /// <param name="token"></param>
    /// <remarks>Exception will be thrown for errors. </remarks>
    public System.Threading.Tasks.Task InstallApp(Stream manifestStream, CancellationToken token = default)
    {
        EwsUtilities.ValidateParam(manifestStream);

        return InternalInstallApp(manifestStream, null, null, false, token);
    }

    /// <summary>
    ///     Install App.
    /// </summary>
    /// <param name="manifestStream">
    ///     The manifest's plain text XML stream.
    ///     Notice: Stream has state. If you want this function read from the expected position of the stream,
    ///     please make sure set read position by manifestStream.Position = expectedPosition.
    ///     Be aware read manifestStream.Lengh puts stream's Position at stream end.
    ///     If you retrieve manifestStream.Lengh before call this function, nothing will be read.
    ///     When this function succeeds, manifestStream is closed. This is by EWS design to
    ///     release resource in timely manner.
    /// </param>
    /// <param name="marketplaceAssetId">The asset id of the addin in marketplace</param>
    /// <param name="marketplaceContentMarket">The target market for content</param>
    /// <param name="sendWelcomeEmail">Whether to send welcome email for the addin</param>
    /// <param name="token"></param>
    /// <returns>True if the app was not already installed. False if it was not installed. Null if it is not a user mailbox.</returns>
    /// <remarks>Exception will be thrown for errors. </remarks>
    internal async Task<bool?> InternalInstallApp(
        Stream manifestStream,
        string? marketplaceAssetId,
        string? marketplaceContentMarket,
        bool sendWelcomeEmail,
        CancellationToken token
    )
    {
        EwsUtilities.ValidateParam(manifestStream);

        var request = new InstallAppRequest(
            this,
            manifestStream,
            marketplaceAssetId,
            marketplaceContentMarket,
            sendWelcomeEmail
        );

        var response = await request.Execute(token).ConfigureAwait(false);

        return response.WasFirstInstall;
    }

    /// <summary>
    ///     Uninstall app.
    /// </summary>
    /// <param name="id">App ID</param>
    /// <param name="token"></param>
    /// <remarks>Exception will be thrown for errors. </remarks>
    public System.Threading.Tasks.Task UninstallApp(string id, CancellationToken token = default)
    {
        EwsUtilities.ValidateParam(id);

        var request = new UninstallAppRequest(this, id);

        return request.Execute(token);
    }

    /// <summary>
    ///     Disable App.
    /// </summary>
    /// <param name="id">App ID</param>
    /// <param name="disableReason">Disable reason</param>
    /// <param name="token"></param>
    /// <remarks>Exception will be thrown for errors. </remarks>
    public System.Threading.Tasks.Task DisableApp(
        string id,
        DisableReasonType disableReason,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(id);
        EwsUtilities.ValidateParam(disableReason);

        var request = new DisableAppRequest(this, id, disableReason);

        return request.Execute(token);
    }

    /// <summary>
    ///     Sets the consent state of an extension.
    /// </summary>
    /// <param name="id">Extension id.</param>
    /// <param name="state">Sets the consent state of an extension.</param>
    /// <param name="token"></param>
    /// <remarks>Exception will be thrown for errors. </remarks>
    public System.Threading.Tasks.Task RegisterConsent(string id, ConsentState state, CancellationToken token = default)
    {
        EwsUtilities.ValidateParam(id);
        EwsUtilities.ValidateParam(state);

        var request = new RegisterConsentRequest(this, id, state);

        return request.Execute(token);
    }

    /// <summary>
    ///     Get App Marketplace Url.
    /// </summary>
    /// <remarks>Exception will be thrown for errors. </remarks>
    public Task<string> GetAppMarketplaceUrl()
    {
        return GetAppMarketplaceUrl(null, null);
    }

    /// <summary>
    ///     Get App Marketplace Url.  Works with Exchange 2013 SP1 or later EWS.
    /// </summary>
    /// <param name="apiVersionSupported">The api version supported by the client.</param>
    /// <param name="schemaVersionSupported">The schema version supported by the client.</param>
    /// <param name="token"></param>
    /// <remarks>Exception will be thrown for errors. </remarks>
    public async Task<string> GetAppMarketplaceUrl(
        string? apiVersionSupported,
        string? schemaVersionSupported,
        CancellationToken token = default
    )
    {
        var request = new GetAppMarketplaceUrlRequest(this)
        {
            ApiVersionSupported = apiVersionSupported,
            SchemaVersionSupported = schemaVersionSupported,
        };

        var response = await request.Execute(token).ConfigureAwait(false);
        return response.AppMarketplaceUrl;
    }

    /// <summary>
    ///     Get the client extension data. This method is used in server-to-server calls to retrieve ORG extensions for
    ///     admin powershell/UMC access and user's powershell/UMC access as well as user's activation for OWA/Outlook.
    ///     This is expected to never be used or called directly from user client.
    /// </summary>
    /// <param name="requestedExtensionIds">An array of requested extension IDs to return.</param>
    /// <param name="shouldReturnEnabledOnly">
    ///     Whether enabled extension only should be returned, e.g. for user's
    ///     OWA/Outlook activation scenario.
    /// </param>
    /// <param name="isUserScope">Whether it's called from admin or user scope</param>
    /// <param name="userId">
    ///     Specifies optional (if called with user scope) user identity. This will allow to do proper
    ///     filtering in cases where admin installs an extension for specific users only
    /// </param>
    /// <param name="userEnabledExtensionIds">
    ///     Optional list of org extension IDs which user enabled. This is necessary for
    ///     proper result filtering on the server end. E.g. if admin installed N extensions but didn't enable them, it does not
    ///     make sense to return manifests for those which user never enabled either. Used only when asked
    ///     for enabled extension only (activation scenario).
    /// </param>
    /// <param name="userDisabledExtensionIds">
    ///     Optional list of org extension IDs which user disabled. This is necessary for
    ///     proper result filtering on the server end. E.g. if admin installed N optional extensions and enabled them, it does
    ///     not make sense to retrieve manifests for extensions which user disabled for him or herself. Used only when asked
    ///     for enabled extension only (activation scenario).
    /// </param>
    /// <param name="isDebug">
    ///     Optional flag to indicate whether it is debug mode.
    ///     If it is, org master table in arbitration mailbox will be returned for debugging purpose.
    /// </param>
    /// <param name="token"></param>
    /// <returns>Collection of ClientExtension objects</returns>
    public Task<GetClientExtensionResponse> GetClientExtension(
        StringList requestedExtensionIds,
        bool shouldReturnEnabledOnly,
        bool isUserScope,
        string userId,
        StringList userEnabledExtensionIds,
        StringList userDisabledExtensionIds,
        bool isDebug,
        CancellationToken token = default
    )
    {
        var request = new GetClientExtensionRequest(
            this,
            requestedExtensionIds,
            shouldReturnEnabledOnly,
            isUserScope,
            userId,
            userEnabledExtensionIds,
            userDisabledExtensionIds,
            isDebug
        );

        return request.Execute(token);
    }

    /// <summary>
    ///     Get the OME (i.e. Office Message Encryption) configuration data. This method is used in server-to-server calls to
    ///     retrieve OME configuration
    /// </summary>
    /// <returns>OME Configuration response object</returns>
    public Task<GetOMEConfigurationResponse> GetOMEConfiguration(CancellationToken token = default)
    {
        var request = new GetOMEConfigurationRequest(this);

        return request.Execute(token);
    }

    /// <summary>
    ///     Set the OME (i.e. Office Message Encryption) configuration data. This method is used in server-to-server calls to
    ///     set encryption configuration
    /// </summary>
    /// <param name="xml">The xml</param>
    /// <param name="token"></param>
    public System.Threading.Tasks.Task SetOMEConfiguration(string xml, CancellationToken token = default)
    {
        var request = new SetOMEConfigurationRequest(this, xml);

        return request.Execute(token);
    }

    /// <summary>
    ///     Set the client extension data. This method is used in server-to-server calls to install/uninstall/configure ORG
    ///     extensions to support admin's management of ORG extensions via powershell/UMC.
    /// </summary>
    /// <param name="actions">List of actions to execute.</param>
    /// <param name="token"></param>
    public System.Threading.Tasks.Task SetClientExtension(
        List<SetClientExtensionAction> actions,
        CancellationToken token = default
    )
    {
        var request = new SetClientExtensionRequest(this, actions);

        return request.ExecuteAsync(token);
    }

    #endregion


    #region Groups

    /// <summary>
    ///     Gets the list of unified groups associated with the user
    /// </summary>
    /// <param name="requestedUnifiedGroupsSets">The Requested Unified Groups Sets</param>
    /// <param name="userSmtpAddress">The smtp address of accessing user.</param>
    /// <param name="token"></param>
    /// <returns>UserUnified groups.</returns>
    public Task<Collection<UnifiedGroupsSet>> GetUserUnifiedGroups(
        IEnumerable<RequestedUnifiedGroupsSet> requestedUnifiedGroupsSets,
        string userSmtpAddress,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(requestedUnifiedGroupsSets);
        EwsUtilities.ValidateParam(userSmtpAddress);

        return GetUserUnifiedGroupsInternal(requestedUnifiedGroupsSets, userSmtpAddress, token);
    }

    /// <summary>
    ///     Gets the list of unified groups associated with the user
    /// </summary>
    /// <param name="requestedUnifiedGroupsSets">The Requested Unified Groups Sets</param>
    /// <param name="token"></param>
    /// <returns>UserUnified groups.</returns>
    public Task<Collection<UnifiedGroupsSet>> GetUserUnifiedGroups(
        IEnumerable<RequestedUnifiedGroupsSet> requestedUnifiedGroupsSets,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(requestedUnifiedGroupsSets);
        return GetUserUnifiedGroupsInternal(requestedUnifiedGroupsSets, null, token);
    }

    /// <summary>
    ///     Gets the list of unified groups associated with the user
    /// </summary>
    /// <param name="requestedUnifiedGroupsSets">The Requested Unified Groups Sets</param>
    /// <param name="userSmtpAddress">The smtp address of accessing user.</param>
    /// <param name="token"></param>
    /// <returns>UserUnified groups.</returns>
    private async Task<Collection<UnifiedGroupsSet>> GetUserUnifiedGroupsInternal(
        IEnumerable<RequestedUnifiedGroupsSet>? requestedUnifiedGroupsSets,
        string? userSmtpAddress,
        CancellationToken token
    )
    {
        var request = new GetUserUnifiedGroupsRequest(this);

        if (!string.IsNullOrEmpty(userSmtpAddress))
        {
            request.UserSmtpAddress = userSmtpAddress;
        }

        if (requestedUnifiedGroupsSets != null)
        {
            request.RequestedUnifiedGroupsSets = requestedUnifiedGroupsSets;
        }

        var response = await request.Execute(token).ConfigureAwait(false);
        return response.GroupsSets;
    }

    /// <summary>
    ///     Gets the UnifiedGroupsUnseenCount for the group specfied
    /// </summary>
    /// <param name="groupMailboxSmtpAddress">The smtpaddress of group for which unseendata is desired</param>
    /// <param name="lastVisitedTimeUtc">The LastVisitedTimeUtc of group for which unseendata is desired</param>
    /// <param name="token"></param>
    /// <returns>UnifiedGroupsUnseenCount</returns>
    public async Task<int> GetUnifiedGroupUnseenCount(
        string groupMailboxSmtpAddress,
        DateTime lastVisitedTimeUtc,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(groupMailboxSmtpAddress);

        var request = new GetUnifiedGroupUnseenCountRequest(
            this,
            lastVisitedTimeUtc,
            UnifiedGroupIdentityType.SmtpAddress,
            groupMailboxSmtpAddress
        )
        {
            AnchorMailbox = groupMailboxSmtpAddress,
        };

        var response = await request.Execute(token).ConfigureAwait(false);
        return response.UnseenCount;
    }

    /// <summary>
    ///     Sets the LastVisitedTime for the group specfied
    /// </summary>
    /// <param name="groupMailboxSmtpAddress">The smtpaddress of group for which unseendata is desired</param>
    /// <param name="lastVisitedTimeUtc">The LastVisitedTimeUtc of group for which unseendata is desired</param>
    /// <param name="token"></param>
    public System.Threading.Tasks.Task SetUnifiedGroupLastVisitedTime(
        string groupMailboxSmtpAddress,
        DateTime lastVisitedTimeUtc,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(groupMailboxSmtpAddress);

        var request = new SetUnifiedGroupLastVisitedTimeRequest(
            this,
            lastVisitedTimeUtc,
            UnifiedGroupIdentityType.SmtpAddress,
            groupMailboxSmtpAddress
        );

        return request.Execute(token);
    }

    #endregion


    #region Validation

    /// <summary>
    ///     Validates this instance.
    /// </summary>
    internal override void Validate()
    {
        base.Validate();

        if (Url == null)
        {
            throw new ServiceLocalException(Strings.ServiceUrlMustBeSet);
        }

        if (PrivilegedUserId != null && ImpersonatedUserId != null)
        {
            throw new ServiceLocalException(Strings.CannotSetBothImpersonatedAndPrivilegedUser);
        }

        // only one of PrivilegedUserId|ImpersonatedUserId|ManagementRoles can be set.
    }

    /// <summary>
    ///     Validates a new-style version string.
    ///     This validation is not as strict as server-side validation.
    /// </summary>
    /// <param name="version"> the version string </param>
    /// <remarks>
    ///     The target version string has a required part and an optional part.
    ///     The required part is two integers separated by a dot, major.minor
    ///     The optional part is a minimum required version, minimum=major.minor
    ///     Examples:
    ///     X-EWS-TargetVersion: 2.4
    ///     X-EWS_TargetVersion: 2.9; minimum=2.4
    /// </remarks>
    internal static void ValidateTargetVersion(string version)
    {
        const char parameterSeparator = ';';
        const string legacyVersionPrefix = "Exchange20";
        const char parameterValueSeparator = '=';
        const string parameterName = "minimum";

        if (string.IsNullOrEmpty(version))
        {
            throw new ArgumentException("Target version must not be empty.", nameof(version));
        }

        var parts = version.Trim().Split(parameterSeparator);
        switch (parts.Length)
        {
            case 1:
            {
                // Validate the header value. We allow X.Y or Exchange20XX.
                var part1 = parts[0].Trim();
                if (parts[0].StartsWith(legacyVersionPrefix))
                {
                    // Close enough; misses corner cases like "Exchange2001". Server will do complete validation.
                }
                else if (IsMajorMinor(part1))
                {
                    // Also close enough; misses corner cases like ".5".
                }
                else
                {
                    throw new ArgumentException("Target version must match X.Y or Exchange20XX.");
                }

                break;
            }
            case 2:
            {
                // Validate the optional minimum version parameter, "minimum=X.Y"
                var part2 = parts[1].Trim();
                var minParts = part2.Split(parameterValueSeparator);
                if (minParts.Length == 2 &&
                    minParts[0].Trim().Equals(parameterName, StringComparison.OrdinalIgnoreCase) &&
                    IsMajorMinor(minParts[1].Trim()))
                {
                    goto case 1;
                }

                throw new ArgumentException("Target version must match X.Y or Exchange20XX.");
            }
            default:
            {
                throw new ArgumentException("Target version should have the form.");
            }
        }
    }

    private static bool IsMajorMinor(string versionPart)
    {
        const char majorMinorSeparator = '.';

        var parts = versionPart.Split(majorMinorSeparator);
        if (parts.Length != 2)
        {
            return false;
        }

        foreach (var s in parts)
        {
            if (s.Any(c => !char.IsDigit(c)))
            {
                return false;
            }
        }

        return true;
    }

    #endregion


    #region Utilities

    /// <summary>
    ///     Creates an HttpWebRequest instance and initializes it with the appropriate parameters,
    ///     based on the configuration of this service object.
    /// </summary>
    /// <param name="methodName">Name of the method.</param>
    /// <returns>
    ///     An initialized instance of HttpWebRequest.
    /// </returns>
    internal async Task<EwsHttpWebRequest> PrepareHttpWebRequest(string methodName)
    {
        var endpoint = AdjustServiceUriFromCredentials(Url);

        var request = await PrepareHttpWebRequestForUrl(endpoint).ConfigureAwait(false);

        if (!string.IsNullOrEmpty(TargetServerVersion))
        {
            request.Headers.TryAddWithoutValidation(TargetServerVersionHeaderName, TargetServerVersion);
        }

        return request;
    }

    /// <summary>
    ///     Processes an HTTP error response.
    /// </summary>
    /// <param name="httpWebResponse">The HTTP web response.</param>
    /// <param name="webException">The web exception.</param>
    internal override void ProcessHttpErrorResponse(
        IEwsHttpWebResponse httpWebResponse,
        EwsHttpClientException webException
    )
    {
        InternalProcessHttpErrorResponse(
            httpWebResponse,
            webException,
            TraceFlags.EwsResponseHttpHeaders,
            TraceFlags.EwsResponse
        );
    }

    /// <summary>
    ///     Adjusts the service URI based on the current type of credentials.
    /// </summary>
    /// <remarks>
    ///     Autodiscover will always return the "plain" EWS endpoint URL but if the client
    ///     is using WindowsLive credentials, ExchangeService needs to use the WS-Security endpoint.
    /// </remarks>
    /// <param name="uri">The URI.</param>
    /// <returns>Adjusted URL.</returns>
    private Uri AdjustServiceUriFromCredentials(Uri uri)
    {
        return Credentials != null ? Credentials.AdjustUrl(uri) : uri;
    }

    #endregion
}
