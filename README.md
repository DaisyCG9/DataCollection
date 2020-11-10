# DataCollection
Software Development Environment:      .NET Framework 4.7.2   VS 2019

To realize automatic service for current manual case assignment analysis.

Have an always-online tool/service that auto detect, identify, record the service ticket that assigned to AAD Auth APAC engineers and the case info with region, support type, support topic and other data entries, add to the targeted Excel Online file stored in OneDrive.

Workflow: 
a)	Get token using ADAL.  
b)	Abstract case numbers from emails. 
c)	Request https://servicedesk.microsoft.com with token and make a request with case number as a request parameter.  
d)	Crawl the contents.  
e)	Write data into excel.

