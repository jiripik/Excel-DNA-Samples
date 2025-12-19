# Excel-DNA Samples

[![.NET Framework](https://img.shields.io/badge/.NET%20Framework-4.5-blue.svg)](https://dotnet.microsoft.com/download/dotnet-framework)
[![Excel-DNA](https://img.shields.io/badge/Excel--DNA-0.32.0-green.svg)](https://excel-dna.net/)
[![PostSharp](https://img.shields.io/badge/PostSharp-3.1.43-orange.svg)](https://www.postsharp.net/)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

A collection of sample projects demonstrating advanced Excel-DNA techniques for building powerful Excel add-ins with C#.

📖 **Companion articles available at [Jiri Pik's Blog](https://jiripik.com/blog)**

---

## 📋 Table of Contents

- [Overview](#overview)
- [Projects](#projects)
  - [01 PostSharpExcelDNA](#01-postsharpexceldna)
- [Prerequisites](#prerequisites)
- [Getting Started](#getting-started)
- [Usage](#usage)
- [Architecture](#architecture)
- [Contributing](#contributing)
- [Resources](#resources)
- [Author](#author)

---

## Overview

This repository showcases practical examples of integrating [Excel-DNA](https://excel-dna.net/) with other powerful .NET libraries to create professional-grade Excel add-ins. Each project demonstrates specific techniques and patterns that can be applied to real-world financial and quantitative applications.

---

## Projects

### 01 PostSharpExcelDNA

**📝 [Read the full article](https://jiripik.com/2017/05/01/postsharp-excel-dna/)**

Demonstrates how to use [PostSharp](https://www.postsharp.net/) Aspect-Oriented Programming (AOP) to simplify Excel-DNA function development by eliminating cross-cutting concerns from your UDF (User-Defined Function) code.

#### Key Features

- **Automatic Array Resizing** — Dynamic resizing of array formulas to fit returned data
- **Cross-Cutting Concerns Separation** — Error handling, validation, and formatting handled via aspects
- **Async Function Support** — Non-blocking Excel functions using `ExcelAsyncUtil`
- **Formula Validation** — Automatic validation of Excel formula length limits

#### Implementation Highlights

| Component | Description |
|-----------|-------------|
| `ExcelFunctions.cs` | Main Excel UDF implementations with clean, focused business logic |
| `ResultResizerAspectAttribute.cs` | PostSharp aspect handling entry/exit/exception scenarios |
| `ExcelResultResizer.cs` | Array formula resizing logic with Excel selection helpers |

#### Example Usage

```csharp
[ExcelFunction(IsMacroType = true)]
[ResultResizerAspect]
public static object TryMe(object parameter)
{
    return ExcelAsyncUtil.Run("TryMe", new[] { parameter }, () =>
    {
        // Your clean business logic here
        var result = new object[100, 100];
        // ... populate result
        return result;
    });
}
```

The `[ResultResizerAspect]` attribute automatically:
- ✅ Validates the calling formula
- ✅ Resizes the output range to fit the returned array
- ✅ Handles exceptions gracefully
- ✅ Manages Excel calculation state

---

## Prerequisites

- **Visual Studio 2015** or later
- **.NET Framework 4.5** or higher
- **Microsoft Excel** (32-bit or 64-bit)
- **NuGet Package Manager**

---

## Getting Started

### 1. Clone the Repository

```bash
git clone https://github.com/jiripik/Excel-DNA-Samples.git
cd Excel-DNA-Samples
```

### 2. Restore NuGet Packages

Open the solution in Visual Studio and restore packages, or run:

```bash
nuget restore "01 PostSharpExcelDNA\PostSharpExcelDNA.sln"
```

### 3. Build the Solution

Build the project in Visual Studio (F6) or via command line:

```bash
msbuild "01 PostSharpExcelDNA\PostSharpExcelDNA.sln" /p:Configuration=Release
```

### 4. Load the Add-in in Excel

After building, locate the generated `.xll` files in the `bin\Debug` or `bin\Release` folder:

- `PostSharpExcelDNA-AddIn-packed.xll` — 32-bit Excel
- `PostSharpExcelDNA-AddIn64-packed.xll` — 64-bit Excel

Open Excel and load the appropriate `.xll` file via **File → Options → Add-ins → Go → Browse**.

---

## Usage

Once the add-in is loaded, you can use the custom functions directly in Excel:

```excel
=TryMe(A1)
```

This function returns a 100×100 array that automatically resizes to fill the appropriate cells.

---

## Architecture

```
01 PostSharpExcelDNA/
├── ExcelFunctions.cs              # Excel UDF definitions
├── ResultResizerAspectAttribute.cs # PostSharp AOP aspect
├── ExcelResultResizer.cs          # Array resizing utilities
├── PostSharpExcelDNA-AddIn.dna    # Excel-DNA configuration
├── packages.config                # NuGet dependencies
└── Properties/
    └── AssemblyInfo.cs            # Assembly metadata
```

### Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| Excel-DNA | 0.32.0 | Excel add-in framework |
| Excel-DNA.Lib | 0.32.0 | Excel-DNA integration library |
| PostSharp | 3.1.43 | Aspect-Oriented Programming framework |

---

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

---

## Resources

- 📚 [Excel-DNA Documentation](https://excel-dna.net/)
- 📚 [PostSharp Documentation](https://doc.postsharp.net/)
- 📚 [Excel-DNA GitHub Repository](https://github.com/Excel-DNA/ExcelDna)
- 💬 [Excel-DNA Google Group](https://groups.google.com/g/exceldna)

---

## Author

**Jiri Pik**

- 🌐 Website: [jiripik.com](https://jiripik.com)
- 📝 Blog: [jiripik.com/blog](https://jiripik.com/blog)
- 💼 GitHub: [@jiripik](https://github.com/jiripik)

---

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

<p align="center">
  <sub>Built with ❤️ using Excel-DNA and PostSharp</sub>
</p>

