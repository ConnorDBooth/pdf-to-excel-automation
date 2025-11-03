import pdfplumber

pdf_path = r"C:\Mould Lab Files\m318047- results_signed.pdf"

def find_outdoor_section(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[1]  # 0-based index, so 1 is the second page
        tables = page.extract_tables()
        
        for table_num, table in enumerate(tables):
            print(f"\nTable {table_num + 1}:")
            transposed = list(zip(*table))
            outdoor_col_index = None
            outdoor_row_index = None

            # Search for "Outdoor" in the transposed table
            for col_idx, col in enumerate(transposed):
                for row_idx, cell in enumerate(col):
                    if cell and cell.strip().lower() == "outdoor":
                        outdoor_col_index = col_idx
                        outdoor_row_index = row_idx
                        break
                if outdoor_col_index is not None:
                    break
            
            if outdoor_col_index is not None:
                #print(f"'Outdoor' found at column {outdoor_col_index+1}, row {outdoor_row_index+1}")
                # Now you can access the whole column or specific rows as needed
                outdoor_column = transposed[outdoor_col_index]
               # print("Outdoor column values:", outdoor_column)
                mold_col_index = outdoor_col_index + 2

                #Create list of mold types and values
                mold_types = list(transposed[0][6:])
                mold_values = list(transposed[mold_col_index][6:])
               # print("Mold types:", mold_types)
                #print("Mold values:", mold_values)
                
                #zip the mold types and values together to create a dictionary
                mold_dict = {}
                for mold_type, value in zip(mold_types, mold_values):
                    if mold_type and mold_type.strip():
                        mold_dict[mold_type.strip()] = value.strip() if value else None
                
                print("This is the dictionary ", mold_dict)
                
                # Fill the dictionary with corresponding values from the outdoor column

                
            else:
                print("'Outdoor' not found in this table.")
                mold_dict = None

    return mold_dict

find_outdoor_section(pdf_path)

