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

using System.Runtime.Serialization;

using JetBrains.Annotations;

namespace Microsoft.Exchange.WebServices.Data;

/// <summary>
///     Represents a remote service exception that has a single response.
/// </summary>
[PublicAPI]
public class ServiceResponseException : ServiceRemoteException
{
    /// <summary>
    ///     Error details Value keys
    /// </summary>
    private const string ExceptionClassKey = "ExceptionClass";

    private const string ExceptionMessageKey = "ExceptionMessage";
    private const string StackTraceKey = "StackTrace";

    /// <summary>
    ///     Gets the ServiceResponse for the exception.
    /// </summary>
    public ServiceResponse Response { get; }

    /// <summary>
    ///     Gets the service error code.
    /// </summary>
    public ServiceError ErrorCode => Response.ErrorCode;

    /// <summary>
    ///     Gets a message that describes the current exception.
    /// </summary>
    /// <returns>The error message that explains the reason for the exception.</returns>
    public override string Message
    {
        get
        {
            // Special case for Internal Server Error. If the server returned
            // stack trace information, include it in the exception message.
            if (Response.ErrorCode == ServiceError.ErrorInternalServerError)
            {
                if (Response.ErrorDetails.TryGetValue(ExceptionClassKey, out var exceptionClass) &&
                    Response.ErrorDetails.TryGetValue(ExceptionMessageKey, out var exceptionMessage) &&
                    Response.ErrorDetails.TryGetValue(StackTraceKey, out var stackTrace))
                {
                    return string.Format(
                        Strings.ServerErrorAndStackTraceDetails,
                        Response.ErrorMessage,
                        exceptionClass,
                        exceptionMessage,
                        stackTrace
                    );
                }
            }

            return Response.ErrorMessage;
        }
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ServiceResponseException" /> class.
    /// </summary>
    /// <param name="response">The ServiceResponse when service operation failed remotely.</param>
    internal ServiceResponseException(ServiceResponse response)
    {
        Response = response;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="T:Microsoft.Exchange.WebServices.Data.ServiceResponseException" />
    ///     class with serialized data.
    /// </summary>
    /// <param name="info">The object that holds the serialized object data.</param>
    /// <param name="context">The contextual information about the source or destination.</param>
    [Obsolete(
        "This API supports obsolete formatter-based serialization. It should not be called or extended by application code.",
        DiagnosticId = "SYSLIB0051",
        UrlFormat = "https://aka.ms/dotnet-warnings/{0}"
    )]
    protected ServiceResponseException(SerializationInfo info, StreamingContext context)
        : base(info, context)
    {
        Response = (ServiceResponse)info.GetValue("Response", typeof(ServiceResponse));
    }

    /// <summary>
    ///     Sets the <see cref="T:System.Runtime.Serialization.SerializationInfo" /> object with the parameter name and
    ///     additional exception information.
    /// </summary>
    /// <param name="info">The object that holds the serialized object data. </param>
    /// <param name="context">The contextual information about the source or destination. </param>
    /// <exception cref="T:System.ArgumentNullException">
    ///     The <paramref name="info" /> object is a null reference (Nothing in
    ///     Visual Basic).
    /// </exception>
    [Obsolete(
        "This API supports obsolete formatter-based serialization. It should not be called or extended by application code.",
        DiagnosticId = "SYSLIB0051",
        UrlFormat = "https://aka.ms/dotnet-warnings/{0}"
    )]
    public override void GetObjectData(SerializationInfo info, StreamingContext context)
    {
        EwsUtilities.Assert(info != null, "ServiceResponseException.GetObjectData", "info is null");

        base.GetObjectData(info, context);

        info.AddValue("Response", Response, typeof(ServiceResponse));
    }
}
