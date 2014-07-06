old_stats_parser
================

Takes ABBY ocr software output from analyzing scanned old statistics books and transform it in an excel spreadsheet formatted like a database.

Up to now, it can process only a certain type of argentinian old stats book.

The idea is to expand functionality to be able to process more kinds of stat
books and thus liberate them from their physical trap :)

Excel files
-----------

* abby_file.xlsx - This is the file we want to parse
* abby_parsed.xlsx - This is the output
* test_results - Keep previous abby_parsed.xlsx files for tracking improvement

Installation
------------

The repo is structured like a package, so it can be installed from pip using
github clone url. From command line type:

```
pip install git+https://github.com/abenassi/old_stats_parser.git
```

To upgrade the package if you have already installed it:

```
pip install git+https://github.com/abenassi/old_stats_parser.git --upgrade
```

You could also just download or clone the repo and import the package from
old_stats_parser folder.

```python
import os
os.chdir("C:\Path_where_repo_is")
import old_stats_parser
```

How to use it
-------------

1- You can import abby_file module and call main function with or without
input/ouput file names. If they are not provided, defaults are used.

```python
import old_stats_parser.abby_file as abby_file
abby_file.scrape_abby_file("abby_file.xlsx", "abby_parsed.xlsx")
```

2- You can run abby_file directly. Optionally you can pass parameters for
input/output file names. In windows:

```
cd "C:\Path_where_xl_files_are"
python C:\Path_where_abby_file_is\abby_file.py abby_file.xlsx abby_parsed.xlsx
```