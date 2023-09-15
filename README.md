# Quick introduction

This is a .NET Core 7 port of the [Microsoft.Exchange.WebServices.NETStandard](https://www.nuget.org/packages/Microsoft.Exchange.WebServices.NETStandard) package with
minor modifications.

Here are some tips to take into account:

- Almost all functions involving HTTP requests are now async
- Outdated async Begin/End functions were removed
- NET Standard: LDAP Autodiscovery feature will not work
- NET Standard on Linux: DNS Autodiscovery feature will not work
- `ExchangeService.ServerCertificateValidationCallback` can now be specified to configure SSL validation
- Nuget package can be found here: [https://www.nuget.org/packages/Exchange.WebServices.NETCore/](https://www.nuget.org/packages/Exchange.WebServices.NETCore/)

# Getting Started with the EWS Managed API
[![Gitter](https://badges.gitter.im/JoinChat.svg)](https://gitter.im/OfficeDev/ews-managed-api?utm_source=badge&utm_medium=badge&utm_campaign=pr-badge&utm_content=badge)

The Exchange Web Services (EWS) Managed API provides a managed interface for developing .NET client applications that use EWS.
By using the EWS Managed API, you can access almost all the information stored in an Office 365, Exchange Online, or Exchange Server mailbox. However, this API is in sustaining mode, the recommended access pattern for Office 365 and Exchange online data is [Microsoft Graph](https://graph.microsoft.com)

## Support statement

Starting July 19th 2018, Exchange Web Services (EWS) will no longer receive feature updates. While the service will continue to receive security updates and certain non-security updates, product design and features will remain unchanged. This change also applies to the EWS SDKs for Java and .NET. More information here: https://developer.microsoft.com/en-us/graph/blogs/upcoming-changes-to-exchange-web-services-ews-api-for-office-365/

## Getting started resources

See the following articles to help you get started:

- [Get started with EWS Managed API client applications](http://msdn.microsoft.com/en-us/library/office/dn567668(v=exchg.150).aspx)
- [How to: Reference the EWS Managed API assembly](http://msdn.microsoft.com/en-us/library/office/dn528373(v=exchg.150).aspx)
- [How to: Set the EWS service URL by using the EWS Managed API](http://msdn.microsoft.com/en-us/library/office/dn509511(v=exchg.150).aspx)
- [How to: Communicate with EWS by using the EWS Managed API](http://msdn.microsoft.com/en-us/library/office/dn467891(v=exchg.150).aspx)
- [How to: Trace requests and responses to troubleshoot EWS Managed API applications](http://msdn.microsoft.com/en-us/library/office/dn495632(v=exchg.150).aspx)

## Documentation

Documentation for the EWS Managed API is available in the [Web services](http://msdn.microsoft.com/en-us/library/office/dd877012(v=exchg.150).aspx) node of the [MSDN Library](http://msdn.microsoft.com/en-us/library/ms123401.aspx).
In addition to the getting started links provided, you can find how to topics and code samples for the most frequently used EWS Managed API objects in the [Develop](http://msdn.microsoft.com/en-us/library/office/jj900166(v=exchg.150).aspx) node. All the latest information about the EWS Managed API, EWS, and related web services can be found under the [Explore the EWS Managed API, EWS, and web services in Exchange](http://msdn.microsoft.com/en-us/library/office/jj536567(v=exchg.150).aspx) topic on MSDN.

## Additional resources

- [Exchange 101 code samples](http://code.msdn.microsoft.com/Exchange-2013-101-Code-3c38582c)
- [EWS Managed API reference](http://msdn.microsoft.com/en-us/library/jj220535(v=exchg.80).aspx)

## Community

Exchange has an active developer community that you can turn to when you need help. We recommend using the [Exchange Server Development forum on MSDN](http://social.msdn.microsoft.com/Forums/en-US/home?category=exchangeserver&filter=alltypes&sort=lastpostdesc), or using the [ews] tag on [StackOverflow](http://stackoverflow.com/questions/tagged/ews).


This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
