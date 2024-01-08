# Comment-Extract Script

Extract comment data from docx (Microsoft Word) file(s) into a xlsx (Microsoft Excel) file. Combines the results from multiple docx files into a single xlsx spreadsheet. 

Does not require Microsoft Word or Excel be installed.

## Examples

Example input files can be found under 'tests'. Run the configuration file as-is (after renaming it to config.toml) to run the script on the files in the tests folder to obtain an example output.

## Input

One or more docx files that have comments identified within them.

## Output

An xlsx file that includes:
- Formatted text of the comment from the referenced document text
  - Formatting includes bold, italic, underline, double underline, strikethrough, double strikethrough, subscript, and superscript. All other formatting is discarded (e.g., font type, size, etc.). Double-strikethrough is not supported in Excel, so a combination of strikethrough and red-text is used to denote double-strikethrough text.
  - If the referenced document text contains footnotes or endnotes, the notes are appended to the end of the referenced document text. For each comment, footnotes and endnotes are combined and always begin at 1.
- Filename
- Comment author, initials, date
- Comment bubble text
- Comment number within document and total comment number out of all documents

## How to Use

- Use Python 3.11+
- Clone the repository or download the repository as a zip file and extract.
- Setup virtual environment (Windows-specific):
  - Using terminal (e.g., Command Prompt, PowerShell):
    - Navigate to cloned/extracted folder.
    - Enter:

          python -m venv env
          env\Scripts\activate
          pip install -r requirements.txt
          
- Customize configuration file (config.toml) in text-editer:
  - __IMPORTANT__: Rename config.SAMPLE.toml to config.toml!
  - Enter path to folder containing one or more docx files with comments (absolute or relative path)
- Run the script in the terminal:  

        python __main__.py

- The output file will be placed in the _output_ folder, unless customized in the configuration file.

## Similar Repositories

- [daniel-sloat\comment-response](https://github.com/daniel-sloat/comment-response)
  - Uses the output from this script to produce a grouped comment-response docx document.