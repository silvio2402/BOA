## OAB Files

Microsoft Outlook uses a file called the Offline Address Book (OAB) to store email information about your contacts. It is stored locally on every machine that has Outlook installed. When you type someone's name and it shows their email, this file is where that comes from. 

This is a fork of [byteDJINN/BOA]([byteDJINN/BOA](https://github.com/byteDJINN/BOA)) with some improvements.

### Quickstart

```python
python ./boa.py -o output.csv -f csv udetails.oab
```

### Usage

```
usage: boa.py [-h] [-o OUTPUT_FILE] [-f {json,csv}] input_file

Parse OAB (Offline Address Book) data and export to JSON or CSV.

positional arguments:
  input_file            Path to the input OAB file.

options:
  -h, --help            show this help message and exit
  -o, --output_file OUTPUT_FILE
                        Path to the output file (optional). If not specified, output will be printed to stdout.
  -f, --format {json,csv}
                        Output format: json (default) or csv.
```
