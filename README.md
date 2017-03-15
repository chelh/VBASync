<h1>VBA Sync Tool&#x2001;<img src='http://i.imgur.com/sQAsBy4.png' alt='VBA Syc Tool logo' /></h1>

Microsoft Office VBA code is usually held in binary format, making proper
version control difficult. VBA Sync Tool synchronizes macros between a
VBA-enabled file and a folder, enabling easy version control using any VCS.

<h2>Features</h2>
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

<h2>Using</h2>
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
settings automatically. You can also specify a settings file on
the command-line. I recommend launching VBA Sync Tool from a shortcut,
to avoid having to specify the folder and file each time you need
to synchronize.

<img src='http://i.imgur.com/GrXx2VH.png' alt='VBA Sync Tool after selecting folder and file locations' />

<h2>Contributing</h2>
I appreciate any code contributions, but am especially interested
in issues [#1](https://github.com/chelh/VBASync/issues/1)
(Access database support) and
[#8](https://github.com/chelh/VBASync/issues/8) (translations).

Opening a [new issue](https://github.com/chelh/VBASync/issues/new) is
the best way to contact me, especially if you want to contribute code.

Build using Visual Studio 2017. You must also download
[`VBACompressionCodec.dll`](https://github.com/chelh/VBACompressionCodec/releases)
to the `src` directory, as the `VBASync.Model` project depends on it.

<h2>License</h2>
Copyright © 2017 Chelsea Hughes

You may use this software for any purpose and alter it freely.
You may redistribute it subject to these restrictions:

 1. Don’t misrepresent the software’s origin.
 2. Clearly mark any altered versions, and don’t misrepresent them
    as the original.
 3. Keep this notice intact when you distribute the software’s
    source code.

This software is provided “as-is,” without any express or
implied warranty. In no event will I be held liable for any damages
arising from the use of this software.
