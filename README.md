# EasyExcelFrameworkCore
## A generic excel based Keyword Test Framework

### Overview
Keyword test frameworks have been around since the earliest days of test automation and have covered almost every tool. They have also tended to be proprietary and restricted to a single project or organisation. EasyExcelFramework breaks this mould by being completely open source and generic (being able to use any testing API) as well as being extensible.  So, you can take the core framework and adapt it for any test tool, e.g. [Easy Excel Framework Selenium](https://github.com/BigIan1969/EasyExcelFrameworkSelenium).

The other basic aim, as with all Keyword Frameworks, is to make it as easy to use for boots-on-the-ground testers as possible. Hence most test script development is conducted in Excel, an extremely simple tool that almost all testers have a reasonable familiarity with.  This is not to say that all Testers are Luddites.  Just that some are less technically confident than others, and would be overwhelmed by having to do anything but the most basic things in Visual Studio.

### The Core Framework
EasyExcelFrameworkCore is designed to be the foundation upon which to build a Keyword Testing Framework.  The Core, therefore, only has the base functionality that any Keyword Testing Framework should need (Resource loading, Variable management, If/Switch/Loop handling etc).  It is designed to work with addons to do any actual testing, e.g. [Easy Excel Framework Selenium](https://github.com/BigIan1969/EasyExcelFrameworkSelenium).

### Install fram NuGet
Install-Package EasyExcelFramework

See the [Wiki](https://github.com/BigIan1969/EasyExcelFrameworkCore/wiki) for further info
