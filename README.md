# PstToolkit

A high-performance C# library for reading, creating, and transferring email messages between PST files with advanced filtering and extraction capabilities.

## Overview

PstToolkit is a .NET Core library that provides a comprehensive set of tools for working with PST files without using Outlook Interop or any paid libraries. The toolkit allows direct manipulation of PST binary structure, enabling developers to:

- Open existing PST files for reading and extraction
- Create new PST files from scratch
- Add folders and messages to PST files
- Copy messages between PST files with advanced filtering
- Extract emails to standard .eml format
- Navigate folder hierarchies

## Key Features

- **Pure .NET Implementation**: No dependencies on Outlook or COM interop
- **Cross-Platform**: Works on Windows, macOS, and Linux through .NET Core
- **Advanced Filtering**: Filter messages by date, sender, subject, content, and more
- **High Performance**: Optimized for handling large PST files efficiently
- **Memory-Efficient**: Uses streaming approaches for minimal memory footprint
- **Attachment Handling**: Supports extracting and processing of message attachments
- **Unicode Support**: Full support for Unicode characters in all fields
- **Message Preservation**: Maintains original message properties when copying

## Getting Started

### Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/PstToolkit.git
cd PstToolkit

# Build the solution
dotnet build
```

### Usage Examples

The solution includes a demo application that showcases the main features of the library:

```bash
# List all folders in a PST file
dotnet run --project PstToolkitDemo/PstToolkitDemo.csproj list path/to/file.pst

# List all messages in a specific folder
dotnet run --project PstToolkitDemo/PstToolkitDemo.csproj listmessages path/to/file.pst "Inbox/Subfolder"

# Extract messages to .eml files
dotnet run --project PstToolkitDemo/PstToolkitDemo.csproj extract path/to/file.pst "Inbox/Subfolder" output/directory

# Create a new PST file and add a sample message
dotnet run --project PstToolkitDemo/PstToolkitDemo.csproj create new.pst
dotnet run --project PstToolkitDemo/PstToolkitDemo.csproj message new.pst "Inbox/Test"

# Copy messages between PST files with filtering
dotnet run --project PstToolkitDemo/PstToolkitDemo.csproj filteredcopy source.pst destination.pst "subject:Meeting"
```

## Command Line Reference

The demo application supports the following commands:

```
info <pst_file>                       - Show information about a PST file
list <pst_file>                       - List all folders in a PST file
listmessages <pst_file> <folder>      - List all messages in a folder
copy <src> <dst>                      - Copy messages from source PST to destination PST
filteredcopy <src> <dst> <filter>     - Copy filtered messages from source PST to destination PST
create <pst_file>                     - Create a new PST file
message <pst_file> <folder>           - Add a sample message to the specified folder
extract <pst_file> <folder> <output>  - Extract messages from a folder to .eml files
createfolder <pst_file> <path>        - Create a folder path (e.g. 'Inbox/Subfolder')
```

## Filter Syntax

The library supports a powerful filter syntax for selecting messages:

- `subject:keyword` - Messages with "keyword" in the subject
- `from:address` - Messages from a specific email address
- `to:address` - Messages sent to a specific email address
- `after:2023-01-01` - Messages received after a specific date
- `before:2023-12-31` - Messages received before a specific date
- `hasattachment:true` - Messages with attachments
- `body:text` - Messages containing specific text in the body
- `isread:false` - Unread messages

Filters can be combined using AND (&&) and OR (||) operators, and grouped with parentheses.

## Library Architecture

The library is organized into logical components:

- **Core Classes**: `PstFile`, `PstFolder`, `PstMessage`
- **Utilities**: Binary readers/writers, property context handling
- **Formats**: PST structure constants and node types
- **Exceptions**: Specialized exception types for error handling

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request