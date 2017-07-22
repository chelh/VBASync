# VBA Sync&#x2001;![VBA Sync logo](http://i.imgur.com/sQAsBy4.png)

Microsoft Office VBA code is usually held in binary format, making proper
version control difficult. VBA Sync synchronizes macros between a
VBA-enabled file and a folder, enabling easy version control using Git, SVN,
Mercurial, or any other VCS.

**I no longer maintain this project. I have released it into the public domain.**

[**Download my final release (v2.2.0)**](https://github.com/chelh/VBASync/releases/latest)  
[**Look for a newer version, or advertise your new version**](https://github.com/chelh/VBASync/issues/36)

## Features
VBA Sync works *directly with the Office file,* unlike most
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
  * Compatible with Windows, Linux, and Mac.

## Using
VBA Sync has two modes: **Extract** mode extracts modules
from the file into the folder. You can then commit the extracted files
to version control. **Publish** mode publishes modules from
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
and located in the working directory, VBA Sync will load those
settings automatically. I recommend taking advantage of this and
launching VBA Sync from a shortcut with the working directory overridden,
to avoid having to specify the folder and file each time you need to synchronize.

![VBA Sync after selecting folder and file locations](http://i.imgur.com/GrXx2VH.png)

## Command-line
You can also specify settings on the command-line via switches:

Switch | Meaning
------ | ------
`-x`   | Extract VBA from Office file (default)
`-p`   | Publish VBA to Office file
`-f <file>` | Specify Office file
`-d <dir>` | Specify version-control directory
`-r`   | Do the selected action, then immediately exit (**required** on Linux/Mac)
`-i`   | Ignore empty modules
`-u`   | Search subdirectories of version-control directory
`-a`   | Allow adding new document modules when publishing (expert option)
`-e`   | Allow deleting document modules when publishing (expert option)
`-h <hook>` | If `-p` was specified earlier, set the before-publish hook. Else set the after-extract hook.

Any other parameter passed to VBA Sync will be read and parsed as a session `.ini` file.

## Public domain software
Created 2017 by Chelsea Hughes

Thanks to GitHub user hectorticoli for the French translation.

I release all rights to this work. You may use it for any purpose, and alter
and redistribute it freely. If you use this in another product, credit would
be appreciated but is not required.

This software is provided “as-is,” without any express or implied warranty.
In no event will I or any other contributor be held liable for any damages
arising from the use of this software.
