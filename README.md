# PstToolkit

A C# library for reading, creating, and transferring email messages between PST files without using Outlook Interop or paid libraries.

## Features

- Open and read PST files
- Extract email messages and their metadata
- Create new PST files programmatically
- Copy messages between PST files
- Handle various PST file formats and versions (ANSI and Unicode)
- Manage PST file structure (folders, subfolders)
- No dependency on Outlook Interop
- No use of paid/commercial libraries

## Requirements

- .NET 6.0 or later
- MimeKit (for email format handling)

## Getting Started

### Installation

Clone this repository and build the solution using Visual Studio or the .NET CLI:

```bash
git clone https://github.com/yourusername/PstToolkit.git
cd PstToolkit
dotnet build
