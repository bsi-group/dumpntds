ditjson
========

# Background

`ditjson` is a fork of [dumpntds](https://github.com/bsi-group/dumpntds). Unlike the original tool, the purpose it to generate JSON files in order to help integration with other tools. 

This fork updates the underlying framework to .NET 8.0 and uses NuGet packages rather than deploying dependencies as a part of repository. It is possible to publish as a single-file. The trimmed version's size is around 20MB.

The output is a single-file JSON export. The JSON export is opinionated and ignores null values to minimize the exported JSON file size.

# Usage

## Export JSON

Extract the ntds.dit file from the host and run using the following:
```
ditjson -n path\to\ntds.dit\file
```

Once the process has been completed it will have generated two output files in the application directory:

- ntds.json

# Dependencies

- [Commandline](https://github.com/commandlineparser/commandline)
- [ManagedEsent](https://github.com/microsoft/ManagedEsent)