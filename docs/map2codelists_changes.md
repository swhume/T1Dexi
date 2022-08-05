### identifying details of the CT package used for this study
package_date = "2021-12-17"

*Should the above package date be "2022-06-24" as shown in the Word document?*

### codelist OIDs created from c-codes referenced in the SDTM mapping spreadsheet
codelists = ["CL.C66731", "CL.C74457", "CL.C66790"]

*The number of full codelists was reduced due to changes in the mapping spreadsheet
and the addition of more codelist subsets where the entire codelist is not used*

### codelist subset definitions - codelist OID and c-codes and submission values for each term in the subset

*Additional codelist subsets were added and other special lists, such as a list of lab
codes and the list of units, were merged into the main codelists_subsets list*

### OIDs for codelist subsets for domain - each codelist includes the term for one domain

*Domain codelists were updated based on changes to dataset names, and addition of new domains*

### def create_defined_subsets

*The create_defined_subsets method represents a more general approach to processing
codelist subsets since the lab and units subsets were merged into the main list. The comments
section has changed. The ability to process non-standard terms was added.

*** def main()

*The codelist_subsets list has the units and lab subsets merged into it, so no need to process
those subsets separately. One method processes the codelists subsets.