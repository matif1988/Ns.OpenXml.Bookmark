[![Latest version](https://img.shields.io/nuget/v/ns.openxml.bookmark.svg)](https://www.nuget.org/packages/NS.OpenXml.Bookmark)
[![NuGet](https://img.shields.io/nuget/dt/NS.OpenXml.Bookmark.svg)](https://www.nuget.org/packages/NS.OpenXml.Bookmark)
[![master](https://img.shields.io/azure-devops/build/neosys000/oms/5/master.svg)](https://img.shields.io/azure-devops/build/neosys000/oms/5/master.svg)
[![MyGet](https://img.shields.io/azure-devops/release/neosys000/945dc9e7-47f4-4349-8840-e5f4cffa92e4/2/2.svg)](https://img.shields.io/azure-devops/release/neosys000/945dc9e7-47f4-4349-8840-e5f4cffa92e4/2/2.svg)
[![Build status](https://neosys000.visualstudio.com/OMS/_apis/build/status/NS.OpenXml.Bookmark-CI)](https://neosys000.visualstudio.com/OMS/_build/latest?definitionId=5)

# Open XML Excel Interop

The Open XML Excel Interop is a small .Net library that provides tools for working with Office Word using Open XML SDK. It supports scenarios such as:
- Populating content in Word files using Open XML.
- Searching and replacing content in Word using bookmarks (Text, Image, Table, ...).
- Retrieving list of bookmarks in Word document.

Table of Contents
-----------------

- [Dependencies](#dependencies)
- [Releases](#releases)
  - [Supported platforms](#supported-platforms)
- [If You Have Problems](#if-you-have-problems)
- [Support](#support)

Dependencies
------------
* [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml/)
* [HtmlToOpenXml](https://www.nuget.org/packages/NS.HtmlToOpenXml/)

Releases
--------

The official release NuGet packages for Open XML Word Bookmark are [available on Nuget.org](https://www.nuget.org/packages/NS.OpenXml.Bookmark).

The NuGet package for the latest builds of the Open XML Word Bookmark is available as a custom feed on MyGet. You can trust this package source, since the custom feed is locked and only this project feeds into the source. Stable releases here will be mirrored onto NuGet and will be identical.

Supported platforms
-------------------

This library supports many platforms. There are builds for .NET 4.5, .NET 4.6, and .NET Standard 2.0. The following platforms are currently supported:

|    Platform     | Minimum Version |
|-----------------|-----------------|
| .NET Framework  | 4.5             |
| .NET Core       | 1.0             |


If You Have Problems
--------------------

If you want to report a problem (bug, behavior, build, distribution, feature request, etc...) with this class library built by this repository, please feel free to post a new issue and I will try to help.

Support
-------

This project is open source, it was developed to make the manipulation of bookmarks on Word more user-friendly. I thank you in advance for anyone in the community for the possible improvements of the solution as well as the report of possible bugs allowing me and any stakeholder to lead a continuous improvement of this product


