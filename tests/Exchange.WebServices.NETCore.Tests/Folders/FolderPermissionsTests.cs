namespace Exchange.WebServices.NETCore.Tests.Folders;

public class FolderPermissionsTests : IClassFixture<ExchangeProvider>
{
    private readonly ExchangeProvider _provider;

    public FolderPermissionsTests(ExchangeProvider provider)
    {
        _provider = provider;
    }

    [Fact]
    public async Task GetFolderPermissions()
    {
        var service = _provider.CreateTestService();

        
    }
}
