# VBAOneDrive
A VBA Class for converting URIs to local paths

I developed the VBA Class to work with OneDrive URIs. This code should work with the three different types of OneDrive:

* OneDrive associated with a Microsoft/Live account
* OneDrive associated with an Office 365 account
* Sharepoint associated with an Office 365 account

It is not the most sophisticated program and I hope others find it useful.

## Files

* OneDrive.zip - Excel Workbook with macro to demonstrate useage.
* COneDrive.cls - The OneDrive Class file
* basOneDrive.bas - just a simple sub routine to demonstrate usage.
* basWMCRegistary.bas - a collection of functions to interact with Windows registary.
* basWMCStrings.bas - a collection of utility sub routines and functions to handle strings.

## COneDrive.cls Interfaces
* URI - Read/Write - String - the file URI
* IsURI - Read - Boolean - Returns True if URI is a URI or False if it is a local path
* LocalPath - Read - String - Returns Localpath for a given URI
* OneDriveType - Read - String - Returns "Consumer", "Commercial-Personal" or "Commercial-Sharepoint"

There are a bunch of other properties in the class, but those listed above are the most important. The other properties in the class are used to determine how to parse the URI and convert to a local path depending on which type of OneDrive is passed. They are needed internally and are exposed as properties mostly for informational purposes.
