from openpyxl import load_workbook
import xlsxwriter as XLS
import os.path
import json

# --- since splitting on ", " could also get rid of the space in the Masters codelist
codelists = {"DM.RACE": {"IsNonStandard": ["C74457"], "VLM": "No", "type": ["text"], "whereclause": []},
             "DM.ETHNICITY": {"IsNonStandard": ["C66790"], "VLM": "No", "type": ["text"], "whereclause": []},
             "DX.DXTRT": {"IsNonStandard": ["Yes"], "VLM": "No", "type": ["text"], "whereclause": []},
             "FA.FAORRES": {"IsNonStandard": ["Yes", "Yes", "Yes", "Yes", "Yes"], "VLM": "Yes",
                            "type": ["text", "text", "text", "text", "text"],
                            "whereclause": [{"variable": "FATESDTCD", "comparator": "EQ", "value": "AGE"},
                                            {"variable": "FATESDTCD", "comparator": "EQ", "value": "SICKTODY"},
                                            {"variable": "FATESDTCD", "comparator": "EQ", "value": "INSCHFL"},
                                            {"variable": "FATESDTCD", "comparator": "EQ", "value": "STRESTDY"},
                                            {"variable": "FATESDTCD", "comparator": "EQ", "value": "SLEEPQLT"}
                                        ]
                            },
             "FADX.FAOBJ": {"IsNonStandard": ["Yes", "Yes", "No"], "VLM": "Yes", "type": ["text", "text", "integer"],
                              "whereclause": [{"variable": "FAOBJ", "comparator": "EQ", "value": "INSULIN PUMP OR CLOSED LOOP"},
                                            {"variable": "FAOBJ", "comparator": "EQ", "value": "CGM"},
                                            {"variable": "FAOBJ", "comparator": "EQ", "value": "CGM Use Last Month"}
                                        ]
                              },
             "FADX.DISINSEX": {"IsNonStandard": ["Yes"], "VLM": "No", "type": ["text"], "whereclause": []},
             "FADX.SUSPINS": {"IsNonStandard": ["Yes"], "VLM": "No", "type": ["text"], "whereclause": []},
             "LB.LBTMINT": {"IsNonStandard": ["Yes"], "VLM": "No", "type": ["text"], "whereclause": []},
             "SC.SCORRES": {"IsNonStandard": ["Yes", "Yes"], "VLM": "Yes", "type": ["text", "text"],
                            "whereclause": [{"variable": "SCTESTCD", "comparator": "EQ", "value": "EDULEVEL"},
                                            {"variable": "SCTESTCD", "comparator": "EQ", "value": "INCMLVL"}
                                            ]
                            }
            }

worksheet_skip = ["T1Dexi SDTM Summary", "T1Dexi Tables", "Domains", "Sheet1"]
excel_map_file = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'data', 'SDTM-mapping-spec-02Feb2022.xlsx')
subset_file = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'data', 'cl_subsets.json')
codelist_file = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'data', 'codelist_subsets-test.xlsx')
header = ["OID", "Name", "NCI Codelist Code", "Data Type", "Order", "Term", "NCI Term Code", "Decoded Value",
          "Comment", "IsNonStandard", "StandardOID"]
wc_header = ["OID", "Dataset", "Variable", "Comparator", "Value", "Comment"]
vlm_header = ["OID", "Order", "Dataset", "Variable", "ItemOID", "Where Clause", "Data Type", "Length",
              "Significant Digits", "Format", "Mandatory", "Codelist", "Origin Type", "Origin Source", "Pages",
              "Method", "Predecessor", "Comment"]


def write_subset_file():
    with open(subset_file, 'w') as file_out:
        json.dump(codelists, file_out)


def process_map_sheet(sheet, domain):
    # assumes not more than 6 VLM (max_row=17)
    for col_num, col in enumerate(sheet.iter_cols(min_col=2,max_col=sheet.max_column, min_row=1, max_row=17, values_only=True)):
        # row_dict = {key: "" for key in header}
        if col and "CODELIST" in str(col[0]):
            print(f"processing codelist {col[0]}...")
            codelist, var_name = col[0].split(" - ")
            subset_codes = []
            for row_num, cell in enumerate(col):
                if row_num > 9 and cell:
                    subset_codes.append(cell)
                    # if more than one subset then have VLM and must lookup testcds
            print(f"found {len(subset_codes)} subsets for {var_name} in domain {domain}")
            codelists[domain + "." + var_name]["domain"] = domain
            codelists[domain + "." + var_name]["subset_terms"] = subset_codes


def process_codelists():
    rows = []
    for key, cl in codelists.items():
        domain, variable = key.split(".")
        for idx, data_type in enumerate(cl["type"]):
            row = {key: "" for key in header}
            if cl["IsNonStandard"][idx] == "No":
                continue
            if cl["VLM"] == "No":
                row["OID"] = "CL." + domain + "." + variable
                row["Name"] = "Codelist for " + domain + " " + variable
            else:
                wc_value = cl["whereclause"][idx]["value"].replace(" ", "-")
                row["OID"] = "CL." + domain + "." + variable + "." + wc_value
                row["Name"] = "Codelist for " + domain + " " + variable + " where " + wc_value
            if len(cl["IsNonStandard"][idx]) > 3:
                row["NCI Codelist Code"] = cl["IsNonStandard"][idx]
            else:
                row["NCI Codelist Code"] = ""
            row["Data Type"] = cl["type"][idx]
            order_num = 1
            for term in cl["subset_terms"][idx].split(", "):
                row["Order"] = order_num
                term_row = row.copy()
                term_row["Term"] = term
                term_row["NCI Term Code"] = ""
                term_row["Decoded Value"] = ""
                term_row["Comment"] = ""
                if cl["IsNonStandard"][idx] == "Yes":
                    term_row["IsNonStandard"] = "Yes"
                    term_row["StandardOID"] = ""
                else:
                    term_row["IsNonStandard"] = ""
                    term_row["StandardOID"] = "STD.2"
                rows.append(term_row)
                order_num += 1
    return rows


def find_cl_max_length(codelist):
    max_length = 0
    for term in codelist.split(", "):
        if len(term) > max_length:
            max_length = len(term)
    return max_length


def find_number_length(domain, variable):
    # update to lookup length in variables
    return 3


def find_number_sigdigits(domain, variable):
    # update to lookup significant digits
    return 2


def process_where_clauses():
    rows = []
    for key, cl in codelists.items():
        domain, variable = key.split(".")
        for idx, wc in enumerate(cl["whereclause"]):
            row = {key: "" for key in wc_header}
            if cl["VLM"] == "No":
                continue
            wc_value = wc["value"].replace(" ", "-")
            row["OID"] = "WC." + domain + "." + variable + "." + wc_value
            # row["ItemOID"] = "IT." + domain + "." + variable + "." + wc_value
            # row["OID"] = "VL." + domain + "." + variable
            row["Dataset"] = domain
            row["Variable"] = variable
            row["Comparator"] = wc["comparator"]
            row["Value"] = wc["value"]
            row["Comment"] = ""
            rows.append(row)
    return rows


def process_vlm():
    rows = []
    for key, cl in codelists.items():
        domain, variable = key.split(".")
        for idx, wc in enumerate(cl["whereclause"]):
            row = {key: "" for key in wc_header}
            if cl["VLM"] == "No":
                continue
            wc_value = wc["value"].replace(" ", "-")
            row["Where Clause"] = "WC." + domain + "." + variable + "." + wc_value
            row["ItemOID"] = "IT." + domain + "." + variable + "." + wc_value
            row["OID"] = "VL." + domain + "." + variable
            row["Dataset"] = domain
            row["Variable"] = variable
            row["Data Type"] = cl["type"][idx]
            if row["Data Type"] == "text":
                row["Codelist"] = "CL." + domain + "." + variable + "." + wc_value
                row["Length"] = find_cl_max_length(row["Codelist"])
                row["Significant Digits"] = ""
                row["Format"] = ""
            elif row["Data Type"] == "integer":
                row["Codelist"] = ""
                row["Length"] = find_number_length(domain, variable)
                row["Significant Digits"] = ""
                row["Format"] = ""
            else:
                row["Codelist"] = ""
                row["Length"] = find_number_length(domain, variable)
                row["Significant Digits"] = find_number_sigdigits(domain, variable)
                row["Format"] = str(find_number_length(domain, variable)) + "." + str(row["Significant Digits"])
            # how determine mandatory for VLM?
            row["Mandatory"] = "No"
            row["Order"] = idx + 1
            row["Origin Type"] = ""
            row["Origin Source"] = ""
            row["Pages"] = ""
            row["Method"] = ""
            row["Predecessor"] = ""
            row["Comment"] = ""
            rows.append(row)
    return rows


def write_codelist_to_xls(worksheet, rows, header, row_nbr=0):
    for row in rows:
        row_nbr += 1
        for c, col_name in enumerate(header):
            worksheet.write(row_nbr, c, row[col_name])
    return row_nbr


def write_header_row(worksheet, header, header_format):
    for c, header_label in enumerate(header):
        worksheet.write(0, c, header_label, header_format)


def load_variables():
    return {}


def main():
    def_workbook = XLS.Workbook(codelist_file, {"strings_to_numbers": False})
    header_format = def_workbook.add_format({"bold": True, "bg_color": "#CCFFFF", "border": True, "border_color": "black"})
    map_workbook = load_workbook(filename=excel_map_file, read_only=False, data_only=True)
    for sheet in map_workbook.worksheets:
        if sheet.title not in worksheet_skip:
            print(f"processing {sheet.title}...")
            domain = sheet.title.split()
            process_map_sheet(sheet, domain[0])
    rows = process_codelists()
    write_subset_file()

    # create codelist subset worksheet
    worksheet = def_workbook.add_worksheet("codelists")
    write_header_row(worksheet, header, header_format)
    write_codelist_to_xls(worksheet, rows, header)

    # create where clause worksheet
    rows = process_where_clauses()
    wc_worksheet = def_workbook.add_worksheet("whereclauses")
    write_header_row(wc_worksheet, wc_header, header_format)
    write_codelist_to_xls(wc_worksheet, rows, wc_header)

    # create VLM
    variables = load_variables()
    rows = process_vlm()
    vlm_worksheet = def_workbook.add_worksheet("valuelevel")
    write_header_row(vlm_worksheet, vlm_header, header_format)
    write_codelist_to_xls(vlm_worksheet, rows, vlm_header)

    def_workbook.close()


if __name__ == '__main__':
    main()
