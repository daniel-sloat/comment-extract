# Comment-Extract Script

Extract comment data from docx (Microsoft Word) file(s) into a xlsx (Microsoft Excel) file. Combines the results from multiple docx files into a single xlsx spreadsheet. 

Does not require Microsoft Word or Excel be installed.

## Examples

Example input files can be found under 'tests'. Example output can be found in 'output'.

## Input

One or more docx files that have comments identified within them.

## Output

An xlsx file that includes:
- Formatted text of the comment from the referenced document text
  - Formatting includes bold, italic, underline, double underline, strikethrough, double strikethrough, subscript, and superscript. All other formatting is discarded (e.g., font type, size, etc.). Double-strikethrough is not supported in Excel, so a combination of strikethrough and red-text is used to denote double-strikethrough text.
  - If the referenced document text contains footnote or endnote references, the references are appended to the end of the referenced document text. Footnotes and endnotes are combined, and for each comment, always begin at 1.
- Filename (parsing filename using one delimiter)
  - Provides parsing option to split filename into two (e.g., a document number and document code or author)
- Comment author, initials, date
- Comment bubble text 
  - Provides parsing option to split comment bubble text into two (default assumes the first part is a broad grouping (heading 1) and the second part replaces the parsed document code from the filename)
- Comment number within document and total comment number out of all documents

## How to Use

- Use Python 3.10+
- Clone the repository or download the repository as a zip file and extract.
- Setup virtual environment (Windows-specific):
  - Using terminal (e.g., Command Prompt, PowerShell):
    - Navigate to cloned/extracted folder.
    - Enter:

          python3.10 -m venv env
          env\Scripts\activate
          pip install -r requirements.txt
          
- Customize configuration file (config.toml) in text-editer:
  - IMPORTANT: Rename config.SAMPLE.toml to config.toml!
  - Enter path to folder containing one or more docx files with comments (absolute or relative path)
  - Enter delimiters (for filepath and comment bubble text, if using)
  - If certain formatting should be ignored, specify them. All formatting can be turned off, if desired.
- Run the script in the terminal:  

        python comment_extract.py

- The output file will be placed in the output folder, unless customized in the configuration file.

## Similar Repositories

- [daniel-sloat\comment-response](https://github.com/daniel-sloat/comment-response)
  - Uses the output from this script to produce a grouped comment-response docx document.