# K2Documentation.Samples.Extensions.InlineFunctions
This sample project is an example of a custom Inline Function for K2 Studio. In this case a inline function that returns data from an Excel Service. 

For more on custom inline functions, please see the topic https://help.k2.com/onlinehelp/k2five/DevRef/5.3/default.htm#Extend/WF/Custom-Inline-Function.html

Note: This extension applies to legacy components (such as K2 Studio and K2 for Visual Studio), legacy assemblies, legacy services or legacy functionality. If you have upgraded from K2 blackpearl 4.7 to K2 Five, these items may still be available in your environment. These legacy items will not be available in new installations of K2 Five. These legacy components may also not be available, supported, or behave as described here, in future updates or versions of K2.

## Prerequisites
The sample code has the following dependencies: 
* .NET Assemblies (both assemblies are included with K2 client-side tools install and are also included in the project's References folder)
  * SourceCode.Framework.dll

## Getting started
* Use this project to learn how to extend the K2 platform by creating a custom inline function for K2 Studio.  
* This specific project exposes Excel cell values as an Inline Function that can be used in the K2 Studio workflwo designer. 
* This project should compile and run, and you can use the sample Excel files provided in the Resources folder to test the inline function with. 
* Fetch or Pull the appropriate branch of this project for your version of K2. 
* The Master branch is considered the latest, up-to-date version of the sample project. Older versions will be branched. For example, there may be a branch called K2-Five-5.0 that is specific to K2 Five version 5.0. There may be another branch called K2-Five-5.1 that is specific to K2 Five version 5.3. Assume that the master branch is configured for the latest release version of K2 Five. 
* The Visual Studio project contains a folder called "References" where you can find the referenced .NET assemblies or other uncommon assemblies. By default, try to reference these assemblies from your own environment for consistency, but we provide the referenced assemblies as a convenience in case you are not able to locate or use the referenced assemblies in your environment. 
* The Visual Studio project contains a folder called "Resources". This folder contains addiitonal resources that may be required to use the same code, such as K2 deployment packages, sample files, SQL scripts and so on. 
   
## License
This project is licensed under the MIT license, which can be found in LICENSE.

## Notes
 * The sample code is provided as-is without warranty.
 * These sample code projects are not supported by K2 product support. 
 * The sample code is not necessarily comprehensive for all operations, features or functionality. 
 * We only accept code contributions that are compatible with the MIT license (essentially, MIT and Public Domain).
