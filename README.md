# Mammoth Verbose: Automatic style mapping for Mammoth

This program wraps around the Mammoth python library to convert .docx to HTML, preserving all source style names as classes in the HTML and including original .docx style formatting information as attributes on the output HTML elements.

```
$ python mammoth-verbose.py [--map] [--verbose] -i _filename_
```

## Options

--map: Map source .docx style names to class names in the output HTML. Default is true.

--verbose: Preserve source .docx style formatting as attributes on the output HTML. Default is false.

-i: Input filename. Required.

For example:

```
$ python mammoth-verbose.py --map --verbose -i /Users/hederis/Documents/alice.docx
```

## To-Do

* Add some validation to ensure input filename is docx
* Add some validation to see if styles.xml file exists; if not, fail gracefully.