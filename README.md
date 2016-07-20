<h1>VBA Sync Tool&#x2001;<img src='http://i.imgur.com/sQAsBy4.png' alt='VBA Syc Tool logo' /></h1>

Microsoft Office VBA code is usually held in binary format, making proper version control difficult. VBA Sync Tool synchronizes macros between a VBA-enabled file and a folder, enabling easy version control using any version control system. It has two modes: **Extract**&nbsp;mode extracts modules from the file into the folder. You can then commit the extracted files to version control. **Publish**&nbsp;mode publishes modules from the folder into the file. You should do this after merges. In either mode, you can cherry-pick which modules you want to synchronize.

This tool is superior to many other solutions because it…
  * …Does not add to your VBA code base.
  * …Does not require special Excel settings.
  * …Allows you to use any off-the-shelf VCS.
  * …Allows you to cherry-pick which modules to extract or publish.
  * …Minimizes spurious changes by ignoring case on variable names, making merges easier.
  * …Extracts full code *directly from the Office file,* including several hidden attributes.
  * …Also extracts settings not tied to a particular module, like references.
  * …Generates FRX files compatible with the VBE, but *without* any embedded timestamp.
  * …Allows you to extract or publish a FRM module without necessarily updating its FRX module.
  * …Works with document or worksheet modules in the same way as any other module.
  * …Supports Excel 97-2003, Excel 2007+, Word 97-2003, Word 2007+, PowerPoint 2007+, and Outlook files.
  * …Is completely free and open-source.

<img src='http://i.stack.imgur.com/2etAI.png' alt='VBA Sync Tool after selecting folder and file locations' />
