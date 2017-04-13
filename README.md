# VBA Sync Tool&#x2001;![VBA Sync logo](http://i.imgur.com/sQAsBy4.png)

[**Download latest release**](https://github.com/chelh/VBASync/releases/latest)

Microsoft Office VBA code is usually held in binary format, making proper
version control difficult. VBA Sync Tool synchronizes macros between a
VBA-enabled file and a folder, enabling easy version control using any VCS.

## Features
VBA Sync Tool works *directly with the Office file,* unlike most
other solutions, which use a host application (e.g., Excel) to manipulate
the VBA code. This gives it several advantages:
  * Does not require special Excel settings.
  * Does not add to your VBA code base.
  * Allows you to use any off-the-shelf version control system.
  * Allows you to cherry-pick which modules to extract or publish.
  * Minimizes spurious changes by ignoring case on variable names,
    making merges easier.
  * Extracts full code including several hidden attributes.
  * Also extracts settings not tied to a particular module,
    like references.
  * Generates FRX files compatible with the VBE, but
    *without* any embedded timestamp.
  * Allows you to extract or publish a FRM module without necessarily
    updating its FRX module.
  * Works with document or worksheet modules in the same way
    as any other module.
  * Supports Excel 97-2003, Excel 2007+, Word 97-2003, Word 2007+,
    PowerPoint 2007+, and Outlook files.

## Using
VBA Sync Tool has two modes: **Extract**&nbsp;mode extracts modules
from the file into the folder. You can then commit the extracted files
to version control. **Publish**&nbsp;mode publishes modules from
the folder into the file. You should do this after merges.

After you select a mode, a folder path, and a file path, the tool will
list which modules have changed, with a checkbox next to each. Tick
the checkbox next to each module with changes you'd like to apply.
Double-click an entry to run a diff tool against the old and new files.
(This requires setting up a diff tool under **File**→**Settings**.)
If the underlying files change, click **Refresh**. When you're ready
to synchronize, click **Apply** or **OK**.

You can save and load session settings from the **File** menu. Settings
are saved as `.ini` files. If a settings file is named `VBASync.ini`
and located in the working directory, VBA Sync Tool will load those
settings automatically. I recommend taking advantage of this and
launching VBA Sync Tool from a shortcut with the working directory overridden,
to avoid having to specify the folder and file each time you need to synchronize.

![VBA Sync after selecting folder and file locations](http://i.imgur.com/GrXx2VH.png)

## Command-line
You can also specify settings on the command-line via switches:

Switch | Meaning
------ | ------
`/X`   | Extract VBA from Office file (default)
`/P`   | Publish VBA to Office file
`/F <file>` | Specify Office file
`/D <dir>` | Specify version-control directory
`/R`   | Do the selected action, then immediately exit

Any other parameter passed to VBA Sync Tool will be read and parsed as a session `.ini` file.

## Contributing
I appreciate any code contributions, but am especially interested
in [issue 1](https://github.com/chelh/VBASync/issues/1) (Access support)
and [issue 8](https://github.com/chelh/VBASync/issues/8) (translations).

Opening a [new issue](https://github.com/chelh/VBASync/issues/new) is
the best way to contact me, especially if you want to contribute code.

Before building, download [`VBACompressionCodec.dll`](https://github.com/chelh/VBACompressionCodec/releases)
to the `src` directory. Then build using Visual Studio 2017.

## License
Copyright © 2017 Chelsea Hughes

Thanks to GitHub user hectorticoli for the French translation.

You may use this software for any purpose and alter it freely.
You may redistribute it subject to these restrictions:

 1. Don’t misrepresent the software’s origin.
 2. Clearly mark any altered versions, and don’t misrepresent them
    as the original.
 3. Keep this notice intact when you distribute the software’s
    source code.

This software is provided “as-is,” without any express or
implied warranty. In no event will I or any other contributor
be held liable for any damages arising from the use of this software.
