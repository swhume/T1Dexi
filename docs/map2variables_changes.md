### for variables like STUDYID and USUBJID - just define these variables once
common_variables = ["STUDYID", "USUBJID", "SPDEVID"]

*Certain variable definitions are re-used across many domains and should only be defined once. The
common_variables list above indicates that these variables are defined once, and not for every domain
in which they are used*

### key sequences for datasets
key_sequence = {
    "CM": ["STUDYID", "USUBJID", "CMTRT"],

*The key sequence was sent to me separately, and I added that metadata into the code so that 
variables that are keys can be identified with their order preserved.*

### name and path of the input SDTM mapping spreadsheet and default -i CLI arg value - assumes child data dir
excel_map_file = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'data', 'SDTM-mapping-spec-20220406.xlsx')

*This change reflects a new, more recent version of the SDTM mapping spreadsheet to be used by this program. It
changed from the previous version named 'SDTM-mapping-spec-02Feb2022.xlsx'.*

### class Name

*Updated to process common variables (re-used across domains) correctly by defining them once and
re-using them and to add the key_sequence attribute to variables identified as keys.*

### class DataType

*Updated how the data types are interpreted based on feedback on an earlier version.*



