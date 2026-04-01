generate a README.md for the OpenXmlSdkTs project.  Note that it includes the following features:
- Opening and saving binary documents using JSZip
- Opening and saving flatOpc documents (Useful when writing Word JavaScript/TypeScript add-in applications)
- Opening and saving binary documents that are stored as base64 strings.
- All of the Open XML content types and relationship types are referenced using labels, instead of their long, not very user friendly URIs.  This makes it easy to navigate to the mainDocument part, then from part to part, using methods that are similar to the methods in the dotnet Open-Xml-Sdk.
- Note that this uses, of course, the LINQ to XML for TypeScript library that I finished last week.
- All of the namespaces, element, and attribute names that you will use when querying or modifying the various parts are pre-initialized in static classes, in a similar way to how they are initialized in the Open-Xml-PowerTools library. This makes it convenient to write your code.  Because XNamespace and XName objects are atomized (which means that two XName objects with the name namespace and the same name will be in fact the same object), this gives super-good performance to the code that you write to query and modify markup.

This library, like the LtXmlTs library, is licensed under the MIT license.

Both libraries are consumable from the npmjs, the default package manager for nodejs.

Include a class hierarchy diagram.

Include a Quick Start section

Include an example that
- Loads from a File.
- Modifies the markup in the mainDocument part.  For example, you can capitalize the first word of the document.
- Saves back to a File.

Include anything else that is relevant to an end-user who is deciding whether or not to use this library.
