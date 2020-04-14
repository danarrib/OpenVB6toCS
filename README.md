# Open Visual Basic 6 to C# Migration Tool / Translator

![logo](https://user-images.githubusercontent.com/17026744/79247080-68442000-7e50-11ea-9137-faeec5209107.png)

This is an Open Source and Free Visual Basic 6 to C# Translator. It's intendend to convert Microsoft Visual Basic 6 code to C# code, to facilitate the task of converting the VB6 code to the .NET platform.

The goal is to be capable to convert most of the VB6 code to C#, reducing the amount of refactoring.

Here's some tips that can help to convert:

* It works better for form-less applications (like ActiveX DLLs used as COM+ components)
* Avoid using Modules (.BAS) on your project. Use only Classes.
* Avoid using another referenced components. They are probably the part that you'll need to refactor.

What it already does?
* Fields declarations (`public xyz as String` becomes `public string xyz;`)
* Basic datatype conversion (`String` to `string`, `Long` to `int`, `Integer` to `short`, etc)
* Enum declarations
* Keeps the comments

It's on a VERY EARLY stage of development, so it probably won't be as much useful for you. But feel free to try.
