using System.Diagnostics;

using Microsoft.Exchange.WebServices.Data;

using Task = System.Threading.Tasks.Task;

namespace Exchange.WebServices.NETCore.Tests.Folders;

public class FolderOperationTests : IClassFixture<ExchangeProvider>
{
    private readonly ExchangeProvider _provider;

    public FolderOperationTests(ExchangeProvider provider)
    {
        _provider = provider;
    }

    [Fact]
    public async Task FindFoldersTest()
    {
        var service = _provider.CreateTestService();

        var folders = await service.FindFolders(
            new FolderId(WellKnownFolderName.MsgFolderRoot),
            new FolderView(100, 0)
        );

        Assert.NotEmpty(folders);
    }


    [Fact]
    public async Task FolderBindTest()
    {
        var service = _provider.CreateTestService();

        var folder = await Folder.Bind(service, WellKnownFolderName.ArchiveRoot, PropertySet.FirstClassProperties);

        Assert.NotNull(folder);
    }

    [Fact]
    public async Task SyncFolderTest()
    {
        var service = _provider.CreateTestService();

        var calendarFolder = new FolderId(WellKnownFolderName.Calendar);
        var icc = await service.SyncFolderItems(
            calendarFolder,
            PropertySet.FirstClassProperties,
            null,
            123,
            SyncFolderItemsScope.NormalItems,
            "aaa"
        );
    }
}
