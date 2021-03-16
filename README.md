# WordDiff
C# Based Console Diff-Tool for Word Documents

It depends on Word COM Interop, so Word must be installed on the system.
This EXE-utility is intended to be used in environments that prohibit
powershell script execution.

## Usage

`c:\> WordDiff base_doc derived_doc`

where "base_doc" is the path to the original Word document and "derived_doc"
is the path to a modified/updated document.

Execution may take a couple of seconds depending on documents load time.
A new Word window with changes highlighted appears after the loading delay.

## Building
The utility uses Visual Studio 2019 for building. It requires besides
C# .NET support, the "Office Developer Tools for Visual Studio" component.