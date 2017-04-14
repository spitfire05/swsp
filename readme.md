# swsp - SolidWorks Standard Primitives

[![Build status](https://ci.appveyor.com/api/projects/status/kjh5jo8yupe5lay1?svg=true)](https://ci.appveyor.com/project/spitfire05/swsp)

## Usage

After installation a tab labeled "Standard Primitives"
should appear in the Command Manager, while in the part mode.

![](http://traal.eu/swsp/swsp.JPG)

## What it can do?

**swsp** is a set of few standard primitive macros, aimed at shortening the time required to
create new part in SolidWorks, by automating some repetitive tasks.

Currently, the addin contains following sketch macros:

* L-profile
* U-profile
* T-profile (closed)
* Hexagon
* Circle
* Rectangle
* Circle with revolve axis
* Rectangle with revolve axis

## Requirements

SWSP requires [.NET 3.5 Framework](http://www.microsoft.com/en-us/download/details.aspx?id=22)
to run, and, obviously, [SolidWorks](http://www.solidworks.com/).
The tool was developed & tested with SolidWorks 2013 & 2014.

## Installation

You can download windows installer for [latest release](https://github.com/spitfire05/swsp/releases/latest)
which will install the addin for you.

Remember to restart SolidWorks after installation is complete!

## Changelog

### 1.2.0

* Added sketch macros with revolve axis: rectangle, circle.
* Added *big* version of icons.
* Fixed bug causing UI to not update right after upgrade.

### 1.1.0

* Add Polish localization.
* Add Hexagon sketch macro.
* Add centered rectangle macro.
* Add centered circle sketch macro.
* Improve UI handling switching between 1.x versions should work out-of-the-box.
* User can now check for new version availability from addin interface.

### 1.0.1

* Fix bug causing the addin to silently fail to register with SolidWorks.

### 1.0.0

* Initial release. Can draw L, U and T (closed) profile sketches.
