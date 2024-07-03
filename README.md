# HelperLibrary

`HelperLibrary` is a .NET library that provides a set of helper methods to extend the functionality of common .NET types. This README covers the `EnumExtensions`, `EnumHelpers`, `PowerShellExtensions`, `ProcessExtensions`, and `Helpers` classes, which add useful extensions to the `Enum` type, PowerShell integration, process management, and utility functions.

## EnumExtensions

The `EnumExtensions` class provides a method to retrieve the description of an enumeration value. This is particularly useful when you use the `Description` attribute to annotate your enum values.

### Features

- Retrieve the description of an enum value using the `Description` attribute.
- Return the enum value as a string if no `Description` attribute is found.

### Installation

To use `EnumExtensions`, include the `HelperLibrary` namespace in your project.

### Usage

#### Enum Definition

First, define your enum and annotate its values with the `Description` attribute.

```csharp
using System.ComponentModel;

public enum MyEnum
{
    [Description("This is the first value")]
    FirstValue,
    
    [Description("This is the second value")]
    SecondValue,

    ThirdValue // No description attribute
}
```

#### Retrieving Descriptions

You can now use the `GetDescription` extension method to retrieve the description of an enum value.

```csharp
using HelperLibrary;
using System;

public class Program
{
    public static void Main()
    {
        MyEnum value1 = MyEnum.FirstValue;
        MyEnum value2 = MyEnum.SecondValue;
        MyEnum value3 = MyEnum.ThirdValue;

        Console.WriteLine(value1.GetDescription()); // Output: This is the first value
        Console.WriteLine(value2.GetDescription()); // Output: This is the second value
        Console.WriteLine(value3.GetDescription()); // Output: ThirdValue
    }
}
```

## EnumHelpers

The `EnumHelpers` class provides methods to retrieve an enum value from an integer and to retrieve the description of an enum value given its integer representation.

### Features

- Retrieve an enum value from its integer representation.
- Retrieve the description of an enum value given its integer representation using the `Description` attribute.

### Installation

To use `EnumHelpers`, include the `HelperLibrary` namespace in your project.

### Usage

#### Retrieving Enum Value from Integer

You can use the `GetEnumValue` method to retrieve an enum value from its integer representation.

```csharp
using HelperLibrary;
using System;

public class Program
{
    public static void Main()
    {
        MyEnum enumValue = EnumHelpers.GetEnumValue<MyEnum>(0);
        Console.WriteLine(enumValue); // Output: FirstValue
    }
}
```

#### Retrieving Enum Description from Integer

You can use the `GetEnumDescription` method to retrieve the description of an enum value given its integer representation.

```csharp
using HelperLibrary;
using System;

public class Program
{
    public static void Main()
    {
        string description = EnumHelpers.GetEnumDescription<MyEnum>(0);
        Console.WriteLine(description); // Output: This is the first value
    }
}
```

## PowerShellExtensions

The `PowerShellExtensions` class provides a variety of helper methods to interact with PowerShell commands, manage PCs, and perform Active Directory (AD) operations.

### Features

- Invoke PowerShell commands and clear the commands afterward.
- Retrieve PowerShell version and execution policy.
- Check if a PC is online, restart or shutdown PCs.
- Manage files and directories remotely using PowerShell.
- Interact with Active Directory to retrieve computer and user information.

### Installation

To use `PowerShellExtensions`, include the `HelperLibrary` namespace in your project.

### Usage

#### Retrieving PowerShell Version

You can use the `GetPSVersion` method to retrieve the current PowerShell version.

```csharp
using HelperLibrary;
using System.Management.Automation;

public class Program
{
    public static void Main()
    {
        using (PowerShell ps = PowerShell.Create())
        {
            string version = ps.GetPSVersion();
            Console.WriteLine(version); // Output: (your PowerShell version)
        }
    }
}
```

#### Restarting a PC

You can use the `RestartPC` method to restart a remote PC.

```csharp
using HelperLibrary;
using System.Management.Automation;
using System;

public class Program
{
    public static void Main()
    {
        using (PowerShell ps = PowerShell.Create())
        {
            string pcAddress = "192.168.1.100";
            try
            {
                ps.RestartPC(pcAddress);
                Console.WriteLine($"Successfully restarted {pcAddress}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to restart {pcAddress}: {ex.Message}");
            }
        }
    }
}
```

## ProcessExtensions

The `ProcessExtensions` class provides methods for running and managing processes locally and remotely, including running PowerShell commands, batch files, and installing/uninstalling MSI/EXE files.

### Features

- Run processes locally with specified arguments.
- Run PowerShell commands and scripts.
- Install, repair, and uninstall MSI and EXE files.
- Run processes and commands on remote PCs.

### Installation

To use `ProcessExtensions`, include the `HelperLibrary` namespace in your project.

### Usage

#### Running a Local Process

You can use the `RunProcess` method to run a local process with specified arguments.

```csharp
using HelperLibrary;
using System.Diagnostics;

public class Program
{
    public static void Main()
    {
        using (Process process = new Process())
        {
            var result = process.RunProcess("cmd.exe", "/c echo Hello, World!");
            Console.WriteLine(result.StandardOutput); // Output: Hello, World!
        }
    }
}
```

#### Running a Remote PowerShell Command

You can use the `RunRemotePSCommand` method to run a PowerShell command on a remote PC.

```csharp
using HelperLibrary;
using System.Diagnostics;

public class Program
{
    public static void Main()
    {
        using (Process process = new Process())
        {
            string pcAddress = "192.168.1.100";
            string command = "Get-Process";
            var result = process.RunRemotePSCommand(pcAddress, command);
            Console.WriteLine(result.StandardOutput); // Output: (list of processes)
        }
    }
}
```

## Helpers

The `Helpers` class provides various utility methods for logging, user input validation, process result conversion, and more.

### Features

- Generate log file paths and create log files.
- Add entries to log files.
- Get valid user input with type conversion.
- Convert process results to user sessions.
- Print logged-in computers for a specific user.
- Export process results to CSV files.

### Installation

To use `Helpers`, include the `HelperLibrary` namespace in your project.

### Usage

#### Creating a Log File

You can use the `CreateLogFile` method to create a log file at a specified path.

```csharp
using HelperLibrary;
using System;

public class Program
{
    public static void Main()
    {
        string logPath = Helpers.GetLogPath("MyApp", "Info", "log");
        Helpers.CreateLogFile(logPath);
        Helpers.AddLogEntry(logPath, "Log file created successfully.");
    }
}
```

#### Getting Valid User Input

You can use the `GetValidUserInput` method to get valid user input with type conversion.

```csharp
using HelperLibrary;
using System;

public class Program
{
    public static void Main()
    {
        int userInput = Helpers.GetValidUserInput<int>("Enter a number: ");
        Console.WriteLine($"You entered: {userInput}");
    }
}
```

### Contributing

If you find any issues or have suggestions for improvements, please feel free to create an issue or submit a pull request.

### License

This project is licensed under the MIT License.

