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
///     Represents a folder containing appointments.
/// </summary>
[PublicAPI]
[ServiceObjectDefinition(XmlElementNames.CalendarFolder)]
public class CalendarFolder : Folder
{
    /// <summary>
    ///     Initializes an unsaved local instance of <see cref="CalendarFolder" />. To bind to an existing calendar folder, use
    ///     CalendarFolder.Bind() instead.
    /// </summary>
    /// <param name="service">The ExchangeService object to which the calendar folder will be bound.</param>
    public CalendarFolder(ExchangeService service)
        : base(service)
    {
    }

    /// <summary>
    ///     Binds to an existing calendar folder and loads the specified set of properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to use to bind to the calendar folder.</param>
    /// <param name="id">The Id of the calendar folder to bind to.</param>
    /// <param name="propertySet">The set of properties to load.</param>
    /// <param name="token"></param>
    /// <returns>A CalendarFolder instance representing the calendar folder corresponding to the specified Id.</returns>
    public new static Task<CalendarFolder> Bind(
        ExchangeService service,
        FolderId id,
        PropertySet propertySet,
        CancellationToken token = default
    )
    {
        return service.BindToFolder<CalendarFolder>(id, propertySet, token);
    }

    /// <summary>
    ///     Binds to an existing calendar folder and loads its first class properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to use to bind to the calendar folder.</param>
    /// <param name="id">The Id of the calendar folder to bind to.</param>
    /// <param name="token"></param>
    /// <returns>A CalendarFolder instance representing the calendar folder corresponding to the specified Id.</returns>
    public new static Task<CalendarFolder> Bind(ExchangeService service, FolderId id, CancellationToken token = default)
    {
        return Bind(service, id, PropertySet.FirstClassProperties, token);
    }

    /// <summary>
    ///     Binds to an existing calendar folder and loads the specified set of properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to use to bind to the calendar folder.</param>
    /// <param name="name">The name of the calendar folder to bind to.</param>
    /// <param name="propertySet">The set of properties to load.</param>
    /// <param name="token"></param>
    /// <returns>A CalendarFolder instance representing the calendar folder with the specified name.</returns>
    public new static Task<CalendarFolder> Bind(
        ExchangeService service,
        WellKnownFolderName name,
        PropertySet propertySet,
        CancellationToken token = default
    )
    {
        return Bind(service, new FolderId(name), propertySet, token);
    }

    /// <summary>
    ///     Binds to an existing calendar folder and loads its first class properties.
    ///     Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="service">The service to use to bind to the calendar folder.</param>
    /// <param name="name">The name of the calendar folder to bind to.</param>
    /// <param name="token"></param>
    /// <returns>A CalendarFolder instance representing the calendar folder with the specified name.</returns>
    public new static Task<CalendarFolder> Bind(
        ExchangeService service,
        WellKnownFolderName name,
        CancellationToken token = default
    )
    {
        return Bind(service, new FolderId(name), PropertySet.FirstClassProperties, token);
    }

    /// <summary>
    ///     Obtains a list of appointments by searching the contents of this folder and performing recurrence expansion
    ///     for recurring appointments. Calling this method results in a call to EWS.
    /// </summary>
    /// <param name="view">The view controlling the range of appointments returned.</param>
    /// <param name="token"></param>
    /// <returns>An object representing the results of the search operation.</returns>
    public async Task<FindItemsResults<Appointment>> FindAppointments(
        CalendarView view,
        CancellationToken token = default
    )
    {
        EwsUtilities.ValidateParam(view);

        var responses = await InternalFindItems<Appointment>((SearchFilter)null, view, null, token)
            .ConfigureAwait(false);

        return responses[0].Results;
    }

    /// <summary>
    ///     Gets the minimum required server version.
    /// </summary>
    /// <returns>Earliest Exchange version in which this service object type is supported.</returns>
    internal override ExchangeVersion GetMinimumRequiredServerVersion()
    {
        return ExchangeVersion.Exchange2007_SP1;
    }
}
