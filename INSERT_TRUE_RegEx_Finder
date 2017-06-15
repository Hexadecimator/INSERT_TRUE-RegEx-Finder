'''
@author: logans
'''

import re
import openpyxl as xl

def main():
    # "INSERT_TRUE_Statements.xlsx contains logical equations converted to strings
    # this document contains proprietary data and cannot be linked on this GitHub
    wb_original = xl.load_workbook('INSERT_TRUE_Statements.xlsx')
    sheetnames = wb_original.get_sheet_names()
    INSERT_TRUE_SHEET = wb_original.get_sheet_by_name(sheetnames[0])
    num_Rows = INSERT_TRUE_SHEET.max_row
    
    for rowidx in range(1,num_Rows+1):
        curr_Cell_Coord = "B" + str(rowidx)
        curr_Cell_Content = str(INSERT_TRUE_SHEET[curr_Cell_Coord].value)
        
        text = curr_Cell_Content
        # re.search uses regular expressions on string contained in variable "text" to find any instances of the string pattern.
        # "INSERT_TRUE(*)" (where * can be anything - '.+' means 1 or greater instances can be found-very greedy)
        m = re.search('INSERT_TRUE(.+)\)', text)
        if m:
            found = m.group(0)
            found_Cell_Coord = "C" + str(rowidx)
            
            # the print functions here are to make sure no infinite looping is happening - they are throwaway debugging
            while ')' in found:
                found = found.replace(')','')
                print(") character replaced")
            
            while ", OR(" in found:
                found = found.replace(", OR(", "")
                print (", OR( statement(s) replaced")
            
            if "INSERT_TRUE" in found:
                found = found.replace("INSERT_TRUE","\r\rINSERT_TRUE")
                print("newline character added")
                
            # We add newlines to all INSERT_TRUE statements found, but the first INSERT_TRUE should not have any
            # newlines preceding it since it will be the first thing in the excel cell
            INSERT_TRUE_SHEET[found_Cell_Coord] = found.lstrip()

        else:
            #debug to catch if m string is empty (I think Python will throw an exception if so?)
            print("not found @ cell coord" + curr_Cell_Coord)
    
    # Save the data and close the excel file
    wb_original.save("INSERT_TRUE_Statements.xlsx")
    # Tell user that work is finished and program made it to the end
    print("INSERT_TRUE_Statements.xlsx workbook updated, saved, and closed")
    
main()
