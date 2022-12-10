# VBP (A VBA Package manager)

* [1. The initial plan](#1.%20The%20initial%20plan)
* [2. Dealing with version conflicts](#2.%20Dealing%20with%20version%20conflicts)
* [3. Online Registry](#3.%20Online%20Registry)

## 1. The initial plan

### The interface

Developers will install `VBP` as a VBE addin, or drag and drop the module into DLL. To install modules developers will type commands into the immediate window.

```vb
'From web resource
?VBP.Require("https://raw.githubusercontent.com/sancarn/stdVBA/master/src/stdLambda.cls")

'From relative resource
?VBP.Require("./stdLambda.cls")

'From file resource
?VBP.Require("C:\VBA\stdLambda.cls")

'Fromt registered resource
?VBP.Require("stdVBA/stdLambda")
```

VBP will:

1. If a web resource, download the file to disk
2. Identify the module name
3. Check if the module has already been installed
4. Import the module if not already installed
5. Call `MyModule.InstallDependencies()` or if a class `(new MyClass).InstallDependencies()`

If all modules are successfully installed, `Require()` will return `true`. Else `Require()` will return false and print all errors that occurred.

### Modules which require dependencies

Modules which require dependencies will define a function within the body of the function named `InstallDependencies()`. This should define all dependencies required by the VBA module. If the installation fails the return value should be `false` otherwise it should be `true`. If your class requires any further installation, you can do this in this function. If any other issues occur during install please log these with `VBP.RaiseError()`. For instance:

```vb
class MyIterator
  Function InstallDependencies() as Boolean
    InstallDependencies = true
    InstallDependencies = InstallDependencies AND VBP.Require("https://raw.githubusercontent.com/sancarn/stdVBA/master/src/stdICallable.cls")

    'Example further installation
    downloadSuccess = DownloadFile("https://myResource/myFile.dll", "C:\Temp\myFile.dll")
    If downloadSuccess then 
        InstallDependencies = InstallDependencies AND TRUE
    else
        Call VBP.RaiseError("Failed to download resource 'myFile.dll' to 'C:\Temp\myFile.dll'")
        InstallDependencies = false
    end if
  End Function
end class
```

### Use of Custom DLLs

Not everyone has the ability to utilise the `Temporary` folder. Some people might not use `C` as the standard drive, and others might work in businesses where the C drive is regularly cleared out. It is therefore advised people use a seperate utility for calling DLLs like `UniversalDLLs`.

```vb
class MyClass
  Function InstallDependencies() as Boolean
    InstallDependencies = True
    InstallDependencies = InstallDependencies and VBP.RequireStaticResource("https://myResource/myFile.dll")
    InstallDependencies = InstallDependencies and VBP.Require("stdVBA/stdICallable")
    InstallDependencies = InstallDependencies and VBP.Require("stdVBA/stdDLL")
  End Function
  Sub DoSomething()
    With stdDLL.Create(VBP.getStaticResource("https://myResource/myFile.dll"))
        Dim functionName as stdDLL: set functionName = .CreateFunction("FunctionName", vbLong, vbString, vbString, vbLong)
    End With

    Debug.Print functionName("hello", "world", 10) 
  End Sub
end class
```

### Types of file which can be required

* Require a Workbook Reference
* Require a Worksheet
* Require a Class
* Require a Module
* Require a Userform
* Require a Reference (to a DLL etc.)
* Require a DLL installed at the VBP DLL path

## 2. Dealing with version conflicts

Imagine we have 2 classes, which both include 2 versions of the same class as follows:

stdChrome class:

```vb
'stdChrome.cls
'@remark Uses stdAcc v2.0.0
class stdChrome 
    Public Function InstallDependencies() as Boolean
        InstallDependencies = true
        InstallDependencies = InstallDependencies and VBP.Requires("https://raw.githubusercontent.com/sancarn/stdVBA/master/src/stdAcc.cls")
    End Function

    '... later ...

    Public Sub Usage()
        set oAcc = stdAcc.CreateFromWindow(hwnd).FindFirst(stdLambda.Create("$1.Name like '*address*bar*' and $1.Role = 'ROLE_WINDOW'"))
    End Sub
end class
```

GoogleMaps class:

```vb
'GoogleMaps.cls
'@remark Uses stdAcc v1.0.0
class GoogleMaps
    Public Function InstallDependencies() as Boolean
        InstallDependencies = true
        InstallDependencies = InstallDependencies and VBP.Requires("https://github.com/sancarn/stdVBA/blob/6e106991e9041a3c0075c9d50fd847903011b334/src/WIP/stdAcc.cls")
    End Function

    '... later ...

    Public Sub Usage()
       set oAcc = stdAcc.CreateFromWindow(hwnd).FindFirst("Name=*Address*Bar*&role=ROLE_WINDOW")
    End Sub
end class
```

These 2 modules can't be installed together. There are several problems here:

* In a `require()` routine, during the check to see if the module has already been installed, `stdAcc` will be found so either:
    * We can assume that they are compatible - which will fail at runtime at `FindFirst`.
    * We can assume that they are incompatible and report a warning to rely on the developer to check for and sort out any dependency conflicts manually
    * We can to calculate whether there is a version conflict (somehow??).
        * We can try to fix the conflict by importing the module under a different name. But we would still have to modify all instaces of name `stdAcc` in GoogleMaps class to the new name. This cannot be done very easily automatically, as sometimes it will affect strings (in case of `stdLambda`) and other times it won't in case of `debug.print "Hello world"`.

One of the major difficulties here is calculating a version incompatibility.

* You could rely on a version number in an annotated comment
* We could scan all method parameters and decide if the interface at least is compatible
    * Even if the interface is correct, this doesn't mean the inputs are correct. 
        * Consider a `FindWindow(withTitle,withText)` becomes `FindWindow(withText, optional withTitle)` in a later version. The interfaces are compatible but the inputs are reversed.
    * Doesn't work for param arrays.

Even with incompatibility being calculatable you still have to think how are we actually going to fix the module? A regex replace likely won't do it...

Ultimately, the fatal flaw really is that VBA has 1 global namespaces, and you can't ever create real namespaces which scope classes to a certain portion of code. In a more modern language like C# you can simply import these under seperate namespaces and they will never interact. Not the case here however.

As a result, the decision is to do nothing when it comes to potential version conflicts. We will just report it as a warning and hope the user fixes the modules to work in their scenario. It definitely feels bad to do this to the developer though, so perhaps theres something which can be done into the future.

## 3. Online Registry

In an ideal world VBA libraries will be installable and searchable via an Online Registry. It's possible that we could use [`npm`](https://www.npmjs.com/). Alternatively a github could be hosted with json files linking to various libraries. 