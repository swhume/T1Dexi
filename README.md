# T1DEXI Define-XML v2.1 Generation

## Introduction
This project provides scripts used as part of the process of generating a Define-XML v2.1 output
for the T1DEXI study. These scripts are used to extract content from the SDTM Mapping spreadsheet
and the CDISC Library to complete the odmlib-based Define-XML metadata spreadsheet. Once the odmlib metadata
spreadsheet is complete, the xlsx2define2-1.py odmlib example program is used to generate a valid
Define-XML xml file.

To generate the HTML rendering of the Define-XML using the define2-1.xls stylesheet follow the steps
in [](this blog post once it's posted).

The scripts in this project are intended to be modified as needed to extract content from a mapping spreadsheet.
As the mapping spreadsheet changes or decisions about what to include in the Define-XML change, the scripts
used to extract content and load it into the odmlib metadata spreadsheet will also change.

## Running map2variables
The map2variables.py program extracts the variables from the mapping spreadsheet and generates a 
Define-XML v2.1 metadata worksheet for variables. This worksheet can then be copied into the odmlib metadata
spreadsheet for the study. This also makes it possible to transfer partial sets of variables. This is a command-line
program with optional command-line arguments, as shown in the following example:

`python map2variables -i ./path/to/mapping_spec.xlsx -o ./path/to/variables_ws.xlsx`

The -i argument is the path and filename for the SDTM mapping specification. The -o argument is the odmlib metadata
spreadsheet output path and filename. Both arguments have default settings so are optional.

This application is based on the T1Dexi SDTM mapping spreadsheet and changes to the format of the spreadsheet
or different spreadsheets may not work with this program.

## Running map2codelists
The map2codelists.py program generates the codelist metadata based used to generate codelists in Define-XML v2.1. 
The codelists were stripped from the SDTM mapping spreadsheet and added to this program. This program looks up the 
CDISC CT codelists in the CDISC Library to generate the needed content including all the terms. There are some 
special cases that are also addressed such as the domain abbreviation codelists and codelists subsets where a subset 
of the terms are used. As with variables, by restricting the list of codelists generated you can generate a subset
of the needed codelists. This is a command-line program with optional command-line arguments, as shown in the following
example:

'python map2codelists.py -a e9a7d1b9bf1a4036ae7b25533123456 -o ./path/to/codelists.xlsx'

The -a argument is the CDISC Library API Key. You will need to generate a CDISC Library API Key if you do not
already have one. The -o argument is the odmlib metadata spreadsheet output path and filename. Both arguments 
have default settings so are optional.

This application is based on the T1Dexi SDTM mapping spreadsheet. The codelist identifiers have been extracted from the 
SDTM mapping spreadsheet. If new CDISC CT codelists are added or codelists are removed, then the list will need to be
updated.

## Running map2subsets
Coming soon...

## Generating Define-XML using odmlib
To generate the Define-XML v2.1 file from the odmlib metadata spreadsheet you will need to install odmlib 
and then clone the odmlib_examples repsoitory. It's easiest to install odmlib from PyPI as follows:

`pip install odmlib`

Alternatively, you could in stall odmlib from the source in the [odmlib repository](https://github.com/swhume/odmlib).

Next clone the odmlib_examples to download the example programs, including xlsx2define2-1.py.

`git clone https://github.com/swhume/odmlib_examples.git`

To ensure everything installed correctly, test run xlsx2define2-1 using the following command-line:

`python xls2define.py -e ./data/odmlib-define-metadata.xlsx -d ./data/odmlib-test-define.xml`

You should find the odmlib-test-define.xml in the ./data directory.

You can also request that after a Define-XML file is generated that it be schema validated and do some additional
conformance checks. To do this you'll need to expand the previously used command-line:

`-v -c -e ./data/odmlib-define-metadata.xlsx -d ./data/odmlib-test-define.xml 
-s "./DefineV211/schema/cdisc-define-2.1/define2-1-0.xsd`

## Future Use
For this project, we made every effort to maximize the use of existing content in the mapping spreadsheet.
For future projects, it would be better to design the metadata needed for Define-XML into the process
up front. This will make it possible to develop re-usable templates and ETL code to generate the Define-XML. Such
an automated process will reduce the labor needed to generate the Define-XML as well as improving the overall 
quality of the output. This process could use the existing odmlib Define-XML metadata spreadsheet, but
alternatively a different spreadsheet format could be used and a new odmlib-based application built to 
generate Define-XML from it.

## Limitations
These programs were written to extract data from the mapping spreadsheet. The mapping spreadsheet
was changing during the process, and some content was not ideally organized for Define-XML generation. So,
the programs may need to be adjusted as the mapping spreadsheet or content changes.
