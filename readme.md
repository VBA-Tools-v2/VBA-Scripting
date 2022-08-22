# VBA-Scripting | [RadiusCore](https://radiuscore.co.nz) VBA Tools

__Status__: _Incomplete_

Native VBA implementation of Microsoft Scripting Runtime (scrun.dll). Aims to provide Mac support for this library.

# Development Environment

__Environment__

Recommended development environment is 64-bit Office-365 Excel running on Windows 10. Highly recommended to have [Rubberduck VBA](https://github.com/rubberduck-vba/Rubberduck/) installed to improve the VBE, this is required for Unit Testing.

__Build__

[VBA-Git](https://github.com/VBA-Tools-v2/VBA-Git) is a tool that provides git repository support for Excel/VBA projects. It should be used to build an _.xlam_ file from repository source.


# Usage
__Manual__

To manually add VBA-Scripting to a VBA project, import any of the target classes to the project using standard VBE import procedures.

__VBA-Git__

To easily include VBA-Scripting in any VBA project, use VBA-Git to build the target project, ensuring this repository is listed as a dependency in the project's configuration file. An example is included below, however additional information on how to do this can be found in the [VBA-Git ReadMe](https://github.com/VBA-Tools-v2/VBA-Git/blob/master/readme.md), 

```
"VBA-Scripting": {
    "git": "https://github.com/VBA-Tools-v2/VBA-Scripting",
    "tag": "v0.2.2",
    "key": "{readonly personal access token}"
}
```

# Example

TODO