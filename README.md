# AdaskoTheBeAsT.Interop.COM

## Overview
AdaskoTheBeAsT.Interop.COM is a .NET library designed to facilitate COM interoperability,
particularly focusing using COM objects without registration them in the system
based on manifest files and executing methods of COM DLLs in a Single Threaded Apartment (STA).
This library provides a structured approach to handle COM interactions and activation contexts in .NET applications.


## Features

Creation COM objects registration free based on path to COM DLL and manifest file.
Execution of COM methods in STA.

## Getting Started

To use AdaskoTheBeAsT.Interop.COM in your project, include the namespace in your project file:

```csharp
using AdaskoTheBeAsT.Interop.COM;
```

## Usage

#### Executing COM Methods

Use the Executor class to execute methods of a COM DLL. You can either execute a single COM method or a collection of them.

#### Single COM Method Execution

```csharp
var result = Executor.Execute("comAssemblyPath", "manifestPath", () =>
{
    // Your COM method call
});
```


### Multiple COM Method Execution

```csharp
var comPathDescriptors = new List<ComPathDescriptor>
{
    new ComPathDescriptor("comAssemblyPath1", "manifestPath1"),
    new ComPathDescriptor("comAssemblyPath2", "manifestPath2"),
    // Add more descriptors as needed
};

var result = Executor.Execute(comPathDescriptors, () =>
{
    // Your COM method calls
});
```