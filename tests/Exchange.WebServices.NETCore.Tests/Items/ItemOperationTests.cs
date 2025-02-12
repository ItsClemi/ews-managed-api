using Microsoft.Exchange.WebServices.Data;

using Exception = System.Exception;
using Task = System.Threading.Tasks.Task;

namespace Exchange.WebServices.NETCore.Tests.Items;

public class ItemOperationTests : IClassFixture<ExchangeProvider>
{
    private readonly ExchangeProvider _provider;

    public ItemOperationTests(ExchangeProvider provider)
    {
        _provider = provider;
    }

    [Fact]
    public async Task ItemSearchFilterTest()
    {
        using var service = _provider.CreateTestService();

        _ = await Folder.Bind(service, WellKnownFolderName.Inbox);

        // The search filter to get unread email.
        var filter = new SearchFilter.SearchFilterCollection(
            LogicalOperator.And,
            new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false)
        );
        var view = new ItemView(1);

        var items = await service.FindItems(WellKnownFolderName.Inbox, filter, view);
        Assert.NotEmpty(items);
    }

    [Fact]
    public async Task ItemSuccessionTest()
    {
        using var service = _provider.CreateTestService();

        _ = await Folder.Bind(service, WellKnownFolderName.Inbox);

        // The search filter to get unread email.
        var filter = new SearchFilter.SearchFilterCollection(
            LogicalOperator.And,
            new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false)
        );
        var view = new ItemView(1);

        var items = await service.FindItems(WellKnownFolderName.Inbox, filter, view);
        Assert.NotEmpty(items);

        foreach (var item in items)
        {
            var mailItem = await Item.Bind(service, item.Id, [ItemSchema.MimeContent,]);

            Assert.NotNull(mailItem.MimeContent);
        }
    }

    [Fact]
    public async Task FindItems_Cancelled_ThrowsOperationCancelledException()
    {
        using var service = _provider.CreateTestService();

        _ = await Folder.Bind(service, WellKnownFolderName.Inbox);

        // The search filter to get unread email.
        var filter = new SearchFilter.SearchFilterCollection(
            LogicalOperator.And,
            new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false)
        );
        var view = new ItemView(1);

        var source = new CancellationTokenSource();
        await source.CancelAsync();

        try
        {
            await service.FindItems(WellKnownFolderName.Inbox, filter, view, token: source.Token);
        }
        catch (OperationCanceledException)
        {
            // Do nothing
        }
    }
}
