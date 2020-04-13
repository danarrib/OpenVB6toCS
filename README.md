# Open Visual Basic 6 to C# Migration Tool / Translator

This is an Open Source and Free Visual Basic 6 to C# Translator. It's intendend to convert Microsoft Visual Basic 6 code to C# code, to facilitate the task of converting the VB6 code to the .NET platform.

It's on a VERY EARLY stage, so it probably won't be as much useful for you.

The goal is to be capable to convert most of the VB6 code to C#, reducing the amount of refactoring.

Here's some tips that can help to convert:

* It works better for form-less applications (like ActiveX DLLs used as COM+ components)
* Avoid using Modules (.BAS) on your project. Use only Classes.
* Avoid using another referenced components. They are probably the part that you'll need to refactor.
