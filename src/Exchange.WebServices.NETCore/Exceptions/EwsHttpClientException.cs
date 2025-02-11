namespace Microsoft.Exchange.WebServices.Data;

internal class EwsHttpClientException : Exception
{
    public bool IsProtocolError { get; }
    public HttpResponseMessage? Response { get; }

    public EwsHttpClientException(Exception e)
        : base(e.Message, e)
    {
    }

    public EwsHttpClientException(HttpResponseMessage response)
        : base(response.ReasonPhrase)
    {
        IsProtocolError = true;
        Response = response;
    }
}
