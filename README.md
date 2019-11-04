Octopus Helper
======

Console application for backup the variable set of each project deployment from Octopus Deployment.

Installation
------------
>$ git clone https://github.com/apidesh/OctopusHelper.git<br/>
>$ open OctopusHelper.sln with VS.NET 2019 or other editor that can be compiled the .NET core 3.0 and above<br/>


Contribute
----------
Project require to compile with library via NuGet Package Manager following
1. .NET Core 3.0 and above
2. DotNetCore.NPOI(1.2.2)
3. Octopus.Client(7.3.3)

Argument Description
----------
1. ServerUrl: Octopus Server URL or IP Address and Port
2. ApiKey: Octopus API Key to access the Octopus system
3. ProjectName: Project to be export the configuration and deployment process steps
4. OutputPath: Out put file path


Example Debuging with VS.NET 2019 Usage  (VS.NET 2017 is not support the .NET core 3.0)
----------
1. Open the OctopusHelper.sln file solution
2. Debug Mode
    Go to project Properties -> Debug -> Application arguments
    And put an example parameter: "http://your.octopus-server.com" "API-XYZXXXX" "Project Names" "C:\Projects\OctopusBackup"


Example Console Usage
----------
OctopusHelper.exe "http://your.octopus-server.com" "API-XYZXXXX" "Project Names" "C:\Projects\OctopusBackup"

