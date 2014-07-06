old-stats-parser
================

It takes ABBY (ocr software) output from analyzing scanned old statistics
books and transform it in an excel spreadsheet formatted like a database.

Up to know, it can process only a certain type of argentinian old stats book.

The idea is to expand functionality to be able to process more kinds of stat
books and thus liberate them from their physical trap :)

Excel files
-----------

* abby_file.xlsx  This is the file we want to parse
* abby_parsed.xlsx  This is the output
* test_results  Keep previous abby_parsed.xlsx files for tracking improvement

How to use it
-------------

1. You can import abby_file module and call main function with or without
input/ouput file names. If they are not provided, defaults are used.

```python
import abby_file
abby_file.scrape_abby_file("abby_file.xlsx", "abby_parsed.xlsx")
```

2. You can run abby_file directly. Optionally you can pass paremeters for
input/output file names.

python abby_file.py "abby_file.xlsx" "abby_parsed.xlsx"