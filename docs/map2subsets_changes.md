### codelists are those codelists documented in columns after the variables (CODELIST - VARIABLE NAME)

*This program processes the value level metadata. In many cases, it uses codelists documented in 
columns after the variables. These codelists vary based on the value of one of the variables in the
dataset. The changes clean up the original dictionary as well as adding new value level metadata.
The ValueLists and the WhereClauses are generated from the content in the codelists dictionary.*

### def generate_wc_name

*This method generates the name of the where_clause element.*