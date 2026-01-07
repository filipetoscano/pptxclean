pptxclean
==========================================================================

.NET tool to rewrite PPTX files, removing notes, comments, and hidden slides.


Installation
--------------------------------------------------------------------------

* Download zip from Releases
* Extract binary file to `c:\programs\bin`
* Add `c:\programs\bin` to `PATH`


Usage
--------------------------------------------------------------------------

```bash
Usage: pptxclean [command] [options]

Options:
  -?|-h|--help  Show help information.

Commands:
  check         Validates a PowerPoint file
  get           Retrieves the content of Powerpoint file as JSON
  release       Creates a revision of Powerpoint file: cleans and increments
                revision

Run 'pptxclean [command] -?|-h|--help' for more information about a command.
```


Example
--------------------------------------------------------------------------

```bash
> ls -al
-rwxr-xr-x 1 uname 197609  826 Jan  3 10:34 My Presentation Live.pptx

> pptxclean release "My Presentation Live.pptx"
...

> pptxclean release "My Presentation Live.pptx"
...

> ls -al
-rwxr-xr-x 1 uname 197609  826 Jan  3 10:34 My Presentation Live.pptx
-rwxr-xr-x 1 uname 197609  826 Jan  3 10:34 My Presentation 20260103 Rev1.pptx
-rwxr-xr-x 1 uname 197609  826 Jan  3 10:34 My Presentation 20260103 Rev2.pptx
```

Releasing a presentation does all of the following:
* Scans the directory of the file, to determine the latest revision number
* Creates a copy of the live document: the original will not be touched
* Removes hidden slides
* Removes all comments
* Removes all metadata
* On the first slide
  - Replaces `<DATE>` with the today's date (in `dd MMMM yyyy` format)
  - Replaces `LIVE` with the revision number (in `Rev d` format)
* Renames the filename, adding the date and revision
