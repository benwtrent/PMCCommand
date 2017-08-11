# PMCCommand
Allows commands to be sent to the VisualStudio package management console from the command line.

## Setup

For this to work, you need to make sure that you have VisualStudio installed on the system. The version assumed is VS2015, but it should work for any VisualStudio version.

## Running with Jenkins or someother build server

Make sure that the user that will be doing the action has the appropriate rights. [This StackOverflow Answer](https://stackoverflow.com/questions/1491123/system-unauthorizedaccessexception-retrieving-the-com-class-factory-for-word-in/2560877#2560877) addresses the specific problem. Instead of applying the changes to `Microsoft Word Document` set it for `Microsoft Visual Studio <version>`. 

## Additional Resources if you are interested

- [Get DTE Reference](https://msdn.microsoft.com/en-us/library/68shb4dw.aspx)
- [NuGet Package Management Console Command GUIDs and IDs](https://github.com/mono/nuget/tree/master/src/VsConsole/Console)
- [Visual Studio Commands](https://msdn.microsoft.com/en-us/library/cc826040.aspx)
  - [Helpful MSDN List](https://msdn.microsoft.com/en-us/library/microsoft.visualstudio.vsconstants.aspx), just look for GUIDs in there.
  - Just recursively grep for the GUID/ID in that directory to see the origin of a given GUID or Command ID (`C:\Program Files (x86)\Microsoft Visual Studio 14.0\VSSDK\VisualStudioIntegration\Common\Inc` for my machine)
- [Nuget Issue that spurned this work](https://github.com/NuGet/Home/issues/1512)
- [Using Messagefilter for DTE COM interactions](https://msdn.microsoft.com/en-us/library/ms228772.aspx)
  - [Additional Example of OleMessageFilter](http://dl2.plm.automation.siemens.com/solidedge/api/sesdk_web/OleMessageFilterUsage.html)
- [COM IDs in a Nutshell](https://www.codeproject.com/Articles/1265/COM-IDs-Registry-keys-in-a-nutshell)
