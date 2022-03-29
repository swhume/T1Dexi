import requests
import xlsxwriter as XLS
import os.path
import argparse

"""
map2codelists.py generates the codelist metadata based used to generate codelists in Define-XML v2.1. The codelists were
stripped from the SDTM mapping spreadsheet and added to this program. This program looks up the CDISC CT codelists in
the CDISC Library to generate the needed content including all the terms. There are some special cases that are also
addressed such as the domain abbreviation codelists and codelists subsets where a subset of the terms are used.
Example Cmd-line:
    python map2codelists.py -a e9a7d1b9bf1a4036ae7b25533123456 -o ./path/to/codelists.xlsx
"""

# api key - put your own API key here
library_api_key = "e9a7d1b9bf1a4036ae7b25533a081565"

# identifying details of the CT package used for this study
package_date = "2021-12-17"
package_standard = "sdtmct"

# codelist OIDs created from c-codes referenced in the SDTM mapping spreadsheet
codelists = list(set(["CL.C66728", "CL.C66781", "CL.C66731", "CL.C74457", "CL.C66790", "CL.C66729", "CL.C99079",
                      "CL.C99079", "CL.C102580", "CL.C66789", "CL.C99079", "CL.C66742", "CL.C71113", "CL.C99079",
                      "CL.C74559", "CL.C103330", "CL.C66789", "CL.C99079", "CL.C66741", "CL.C67153", "CL.C71148",
                      "CL.C66770", "CL.C66789", "CL.C74456",  "CL.C99073", "CL.C99079", "CL.C78735", "CL.C71620"]))

# labtests used to create lab test codelist subsets (only include the tests that are used)
lab_test_codelists = {"test_code": "C65047", "test_name": "C67154", "test_codes": ["HBA1C"], "test_names": ["Hemoglobin A1C"]}

# placeholder for the unit codelist which should be replaced by unit subsets
units_codelist = {"codelist": "C71620", "units": [""]}

# OIDs for codelist subsets for domain - each codelist includes the term for one domain
domain_codelists = ["CL.DOMAIN.VS", "CL.DOMAIN.SC", "CL.DOMAIN.QS", "CL.DOMAIN.RP", "CL.DOMAIN.PR", "CL.DOMAIN.ML",
                "CL.DOMAIN.LB", "CL.DOMAIN.FA", "CL.DOMAIN.FACM", "CL.DOMAIN.FADX", "CL.DOMAIN.FAML",
                "CL.DOMAIN.FAPR", "CL.DOMAIN.DX", "CL.DOMAIN.DM", "CL.DOMAIN.DI",
                "CL.DOMAIN.CM"]

# codelist tab column headers that match the odmlib spreadsheet
header = ["OID", "Name", "NCI Codelist Code", "Data Type", "Order", "Term", "NCI Term Code", "Decoded Value",
          "Comment", "IsNonStandard", "StandardOID"]

# the worksheets to skip when processing the SDTM mapping spreadsheet
worksheet_skip = ["T1Dexi Tables", "Domains", "Sheet1"]

# output spreadsheet with just the codelists tab that can be copied to the odmlib spreadsheet - assumes child data dir
excel_define_file = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'data', 'codelists-test.xlsx')


def get_codelist_from_library(endpoint, api_key):
    """
    retrieve a codelist from the CDISC Library using the provided endpoint
    :param endpoint: endpoint string used to create the API call to the Library to retrieve a codelist
    :param api_key: string CDISC Library API key
    :return: json results from Library GET request for a specified codelist
    """
    base_url = "https://library.cdisc.org/api"
    headers = {"Accept": "application/json", "User-Agent": "crawler", "api-key": api_key}
    r = requests.get(base_url + endpoint, headers=headers)
    if r.status_code == 200:
        return r.json()
    else:
        print(f"HTTPError {r.status_code} for url {base_url + endpoint}")


def process_library_codelist(cl_oid, cl):
    """
     given a codelist retrieved from the Library, create the odmlib codelist worksheet rows
    :param cl_oid: the OID for the codelist that is included on each row
    :param cl: the codelist retrieved from the Library
    :return: rows - a list of codelist term rows ready to add to the odmlib codelist worksheet
    """
    rows = []
    for order_num, term in enumerate(cl["terms"]):
        row = {key: "" for key in header}
        row["OID"] = cl_oid
        row["Name"] = cl["name"]
        row["NCI Codelist Code"] = cl["conceptId"]
        row["Data Type"] = "text"
        row["Order"] = order_num + 1
        row["Term"] = term["submissionValue"]
        row["NCI Term Code"] = term["conceptId"]
        row["Decoded Value"] = term["preferredTerm"]
        row["Comment"] = ""
        row["IsNonStandard"] = ""
        row["StandardOID"] = "STD.3"
        rows.append(row)
    return rows


def create_domain_codelist_subsets(worksheet, row_nbr, api_key):
    """ create a codelist subset for each domain used in the study
    :param worksheet: odmlib worksheet object to write to
    :param row_nbr: row number to start appending codelists to the worksheet
    :param api_key: string CDISC Library API key
    :return: row_nbr: integer that indicates where to start appending new codelists
    """
    cl_count = 0
    endpoint = "/mdr/ct/packages/" + package_standard + "-" + package_date + "/codelists/C66734"
    cl = get_codelist_from_library(endpoint, api_key)
    rows = []
    for order_nbr, cl_oid in enumerate(domain_codelists):
        row = {key: "" for key in header}
        row["OID"] = cl_oid
        domain_oid = cl_oid.split(".")
        row["Name"] = "Domain Abbreviation (" + domain_oid[2] + ")"
        row["NCI Codelist Code"] = "C66734"
        row["Data Type"] = "text"
        row["Order"] = 1        # every domain codelist has 1 term
        term = get_domain_term(cl, domain_oid[2])
        row["Term"] = term["Term"]
        row["NCI Term Code"] = term["NCI Term Code"]
        row["Decoded Value"] = term["Decoded Value"]
        row["Comment"] = term["Comment"]
        row["IsNonStandard"] = term["IsNonStandard"]
        row["StandardOID"] = term["StandardOID"]
        rows.append(row)
        cl_count += 1

    row_nbr = write_codelist_to_xls(worksheet, rows, row_nbr)
    print(f"added {cl_count} domain codelist subsets...")
    return row_nbr


def get_domain_term(cl, domain):
    """
    create the codelist term values for a given domain code
    :param cl: the codelist for domains
    :param domain: the 2 letter domain abbreviation to use to find the codelist term details to return
    :return: term dictionary with the details of the domain codelist term
    """
    term = {}
    for domain_term in cl["terms"]:
        if domain_term["submissionValue"] == domain:
            term["Term"] = domain_term["submissionValue"]
            term["NCI Term Code"] = domain_term["conceptId"]
            term["Decoded Value"] = domain_term["preferredTerm"]
            term["Comment"] = ""
            term["IsNonStandard"] = ""
            term["StandardOID"] = "STD.3"
            return term
    term["Term"] = domain
    term["NCI Term Code"] = ""
    term["Decoded Value"] = ""
    term["Comment"] = ""
    term["IsNonStandard"] = "Yes"
    term["StandardOID"] = ""
    return term


def create_lab_codelist_subsets(worksheet, row_nbr, api_key):
    """
    generate codelist subsets for lab test codes and test names
    :param worksheet: the odmlib worksheet to write the codelist subset terms to
    :param row_nbr: the integer row number that indicates where to start appending term rows
    :param api_key: string CDISC Library API key
    :return: integer row number that indicates where to start appending the next terms
    """
    row_nbr = create_codelist_subset(worksheet, row_nbr, lab_test_codelists["test_code"],
                                     lab_test_codelists["test_codes"], api_key)
    row_nbr = create_codelist_subset(worksheet, row_nbr, lab_test_codelists["test_name"],
                                     lab_test_codelists["test_names"], api_key)
    print(f"added lab test code and test name codelist subsets...")
    return row_nbr


def create_codelist_subset(worksheet, row_nbr, c_code, submission_values, api_key):
    """
    generate a codelist subset given the codelist c-code and the list of submission values to include in the subset
    :param worksheet: odmlib worksheet object to write the codelist subset terms to
    :param row_nbr: integer row number that indicates where to begin appending new codelist terms
    :param c_code: string concept code for the codelist to subset
    :param submission_values: list of submission values that will represent the terms included in the subset
    :param api_key: string CDISC Library API key
    :return: integer row number that indicates where to start appending the next terms to the worksheet
    """
    endpoint = "/mdr/ct/packages/" + package_standard + "-" + package_date + "/codelists/" + c_code
    cl = get_codelist_from_library(endpoint, api_key)
    order_nbr = 1
    rows = []
    for term in cl["terms"]:
        if term["submissionValue"] not in submission_values:
            continue
        row = {key: "" for key in header}
        row["OID"] = "CL." + c_code
        row["Name"] = cl["submissionValue"]
        row["NCI Codelist Code"] = c_code
        row["Data Type"] = "text"
        row["Order"] = order_nbr
        row["Term"] = term["submissionValue"]
        row["NCI Term Code"] = term["conceptId"]
        row["Decoded Value"] = term["preferredTerm"]
        row["Comment"] = ""
        row["IsNonStandard"] = ""
        row["StandardOID"] = "STD.1"
        order_nbr += 1
        rows.append(row)
    row_nbr = write_codelist_to_xls(worksheet, rows, row_nbr)
    return row_nbr


def write_codelist_to_xls(worksheet, rows, row_nbr=0):
    """
    write codelist term rows to the odmlib worksheet
    :param worksheet: odmlib worksheet object to write to
    :param rows: a list of codelist row dictionaries to write to the worksheet
    :param row_nbr: integer row number that indicates which worksheet row to begin appending terms
    :return: integer row number that indicates the next row to begin appending terms
    """
    for row in rows:
        row_nbr += 1
        for c, col_name in enumerate(header):
            worksheet.write(row_nbr, c, row[col_name])
    return row_nbr


def write_header_row(worksheet, header_format):
    """
    write the worksheet header column labels to the odmlib worksheet
    :param worksheet: the odmlib worksheet to write the column headers to
    :param header_format: the header format to indicate the style of the header columns
    """
    for c, header_label in enumerate(header):
        worksheet.write(0, c, header_label, header_format)


def set_cmd_line_args():
    """
    get the command-line arguments needed to generate a metadata worksheet for codelists
    :return: return the argparse object with the command-line parameters
    """
    parser = argparse.ArgumentParser()
    parser.add_argument("-a", "--api_key", help="CDISC Library API key",
                        required=False, dest="api_key", default=library_api_key)
    parser.add_argument("-o", "--output_file", help="path and file name of Define-XML v2.1 metadata spreadsheet file",
                        required=False, dest="define_xls", default=excel_define_file)
    args = parser.parse_args()
    return args


def main():
    """
    main driver for creating the odmlib worksheet and writing codelist and associated terms to it
    """
    args = set_cmd_line_args()
    workbook = XLS.Workbook(args.define_xls, {"strings_to_numbers": False})
    header_format = workbook.add_format({"bold": True, "bg_color": "#CCFFFF", "border": True, "border_color": "black"})
    worksheet = workbook.add_worksheet("codelists")
    write_header_row(worksheet, header_format)

    cl_count = 0
    row_nbr = 0
    # add codelists
    for cl_oid in codelists:
        prefix, c_code = cl_oid.split(".")
        endpoint = "/mdr/ct/packages/" + package_standard + "-" + package_date + "/codelists/" + c_code
        cl = get_codelist_from_library(endpoint, args.api_key)
        rows = process_library_codelist(cl_oid, cl)
        row_nbr = write_codelist_to_xls(worksheet, rows, row_nbr)
        cl_count += 1
    print(f"added {cl_count} codelists...")

    # add domain codelists
    row_nbr = create_domain_codelist_subsets(worksheet, row_nbr, args.api_key)

    # add lab test name and lab test code subsets
    row_nbr = create_lab_codelist_subsets(worksheet, row_nbr, args.api_key)

    # add units subset
    # add process a units subset

    workbook.close()


if __name__ == '__main__':
    main()