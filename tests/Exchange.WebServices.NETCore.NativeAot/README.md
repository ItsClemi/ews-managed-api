# Native AOT metadata smoke check

Publish and execute on Windows x64 with the .NET 10 SDK and C++ build tools:

```powershell
dotnet publish Exchange.WebServices.NETCore.NativeAot.csproj -c Release -r win-x64
./bin/Release/net10.0/win-x64/publish/Exchange.WebServices.NETCore.NativeAot.exe
```

Other supported Native AOT hosts can use their own runtime identifier and executable path. 
The check needs no credentials or server. 
It exercises schema field discovery, property names, enum wire names and required server versions after linking. 
Internal entry points are reached through statically analyzable reflection; the test does not root every library field or suppress linker diagnostics.

This checks the metadata paths only, not Native AOT compatibility of every EWS operation or authentication method.


TODO: Expand on this