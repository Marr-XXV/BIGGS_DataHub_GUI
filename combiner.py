"""Combiner: reads staged 'latest/' report files and builds the unified record2025.csv.

Connections:
- Invoked by fetcher after downloads complete; processes per (branch, POS, date).
- Uses pandasbiggs.CSVProcessor for downstream analytics (outside this module).
- The main GUI displays the generated record2025.csv in its viewer.
"""
import os
import shutil
import csv
from pandasbiggs import *
from tqdm import tqdm
import pprint
import re
class Combiner():

    def __init__(self):
        self.parentDir = os.getcwd()
        self.parentDir = self.parentDir.replace("\\","/")
        self.directory = 'latest'
        self.filenames = []
        self.filePaths = ['/latest']
        self.new_branches = []
        self.fNB=open( self.parentDir + "/settings/newBranches.txt","r")
        self.newBranchFile=self.fNB.read()
        for row in self.newBranchFile.splitlines():
            self.new_branches.append(row)
    # * Clean temp directory
    def clean(self, directory):
        """Remove all files and subfolders inside the given directory path."""
        for filename in os.listdir(directory):
            file_path = os.path.join(directory, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print('Failed to delete %s. Reason: %s' % (file_path, e))

    # ** initial function to operate read the names of files in the latest folder and assign them to proper variables separated by "_"
    def generate(self):
        """Scan 'latest' folder, group files by branch/POS/date/filetype, and run GenAppend per date."""
        posFilenames = {}
        filetypes = ["rd1800", "blpr", "discount", "rd5000", "rd5500", "rd5800", "rd5900"]

        folder_path = self.parentDir + self.filePaths[0]

        for self.filename in os.listdir(folder_path):
            name_without_ext = self.filename.rsplit(".", 1)[0]
            parts = name_without_ext.split("_")

            if len(parts) < 5:
                continue  # skip invalid filenames

            pos = parts[2]
            filetype = parts[3]
            branch = parts[1]
            date = parts[4]

            if branch not in posFilenames:
                posFilenames[branch] = {}

            if pos not in posFilenames[branch]:
                posFilenames[branch][pos] = {}

            if date not in posFilenames[branch][pos]:
                posFilenames[branch][pos][date] = {}

            # assign file by filetype under specific date
            posFilenames[branch][pos][date][filetype] = name_without_ext

        # loop through and process
        pprint.pprint(posFilenames)
        # input("Ready to process?")

        for branch, posDict in posFilenames.items():
            self.branch = branch
            for pos, dateDict in posDict.items():
                self.pos = pos
                for date, fileTypes in dateDict.items():
                    self.date = date
                    # only call GenAppend if rd5000 exists for that date
                    self.GenAppend(fileTypes['rd5000'] if 'rd5000' in fileTypes else [], fileTypes)

    # ? processing the the main transaction files and creating copy of the reference for transactions that will be needed in creating the master data.
    def GenAppend(self, filename, fTypes):
        """Build normalized transaction rows for a given branch/POS/date by reading required file types."""
        # ? variables reference for the different file types
        filetypes = ["rd1800", "blpr", "discount", "rd5000", "rd5500", "rd5800", "rd5900"]
        self.proc_files = {}
        item_dict = {}
        item_dept_dict = {}
        disc_dict = {}
        dept_dict = {}
        tnsc_dict = {}
        paym_dict = {}
        blpr_dict = {}
        type_dict = {"D" : "Dine-In",
                    "T" : "Take-Out",
                    "C" : "Delivery"}
        time_dict = ["GY","GY","GY","GY","GY","GY","Breakfast","Breakfast","Breakfast","Breakfast","Breakfast","Lunch","Lunch","Lunch","Lunch","PM Snack","PM Snack","PM Snack","PM Snack","Dinner","Dinner","Dinner","Dinner","GY","GY"]
        e=""
        file=""
        a=1
        b=0
        c=1
        d=""
        end=""

        print("Processing: ", self.branch, " pos: ", self.pos,"Date: ",self.date,)
        file = self.preProc(filename) if filename else []
        # ? read the files in the latest folder based on the file types and assign them to variables
        item_proc = self.preProc(fTypes['rd5500']) if 'rd5500' in fTypes else [] # ? products file
        disc_proc = self.preProc(fTypes['discount']) if 'discount' in fTypes else [] # ? discount file
        dept_proc = self.preProc(fTypes['rd1800']) if 'rd1800' in fTypes else [] #? department file
        tnsc_proc = self.preProc(fTypes['rd5800']) if 'rd5800' in fTypes else [] # ? payments transaction file
        paym_proc = self.preProc(fTypes['rd5900']) if 'rd5900' in fTypes else [] # ? payment file
        # pprint.pprint(tnsc_proc)
        # print(fTypes['rd5800'])
        # hi = input("hi")
        

        blpr_proc = self.preProc(fTypes['blpr']) if 'blpr' in fTypes else []
        
        if item_proc:
            for row in reversed(item_proc.splitlines()):
                line = row.strip().split(",")
                # * Category code line 12 for department
                
                # branch
                if len(line) > 12 and self.branch in self.new_branches:  # must have at least 13 columns
                    item_code = line[0]
                    item_dict[item_code] = {
                        "item_name": line[1],
                        "department_code": line[12]
                    }
                elif len(line) > 3 and self.branch in self.new_branches:
                    item_code = line[0]
                    item_dict[item_code] = {
                        "item_name": line[1],
                        "department_code": line[3]     
                    }
                elif len(line) >= 2:  # make sure we have at least 2 columns
                    item_dict[line[0]] = {
                        "item_name": line[1]
                    }
                    
                elif len(line) == 1 and line[0]:  
                    # optional: handle rows with only one value
                    item_dict[line[0]] = ""
        else:
            item_dict = {}

        if dept_proc:
            for row in reversed(dept_proc.splitlines()):
                line = row.split(",")
                if len(line) >= 2:  # make sure we have at least 2 columns
                    dept_dict[line[0]] = line[1]

                elif len(line) == 1 and line[0]:  
                    # optional: handle rows with only one value
                    dept_dict[line[0]] = ""
        else:
            dept_dict = {}

        if disc_proc:
            for row in reversed(disc_proc.splitlines()):
                line = row.split(",")
                if len(line) >= 2:  # make sure we have at least 2 columns
                    disc_dict[line[0]] = line[1]
                elif len(line) == 1 and line[0]:  
                    # optional: handle rows with only one value
                    disc_dict[line[0]] = ""
        else:
            disc_dict = {}

        if tnsc_proc:
            for row in reversed(tnsc_proc.splitlines()):
                line = row.split(",")
                if(len(line)<11):
                    print("lack 11")
                    print(line)
                elif(len(line)<20):
                    print("lack 20")
                    print(line)
                else:
                    if len(line) >= 20:  # make sure we have at least 2 columns
                        tnsc_dict[line[20]] = line[11]
                    else:  
                        # optional: handle rows with only one value
                        tnsc_dict[line[20]] = ""
        else:
            tnsc_dict = {}

            # print(line)
            # print(line[36])
            # print(line[17])
            # print(tnsc_dict[line[36]])
            # print(tnsc_dict)
            # i = input("huh")
        if paym_proc:
            for row in reversed(paym_proc.splitlines()):
                line = row.split(",")
                if len(line) >= 2:  # make sure we have at least 2 columns
                    paym_dict[line[0]] = line[1]
                else:  
                    # optional: handle rows with only one value
                    paym_dict[line[0]] = ""
        else:
            paym_dict = {}

        if blpr_proc:
            for row in reversed(blpr_proc.splitlines()):
                line = row.split(",")
                #print(str(line) + " " + str(len(line)))
                if(len(line) > 3):
                    #print(len(line[1]))
                    if(len(line[1]) == 11):
                        if(re.search('\"=\"\"(.+?)\"\"\"', line[3])):
                            blpr_dict[re.search('\"=\"\"(.+?)\"\"\"', line[3]).group(1)] = line[1]
                        else:
                            if len(line) >= 4:  # make sure we have at least 2 columns
                                blpr_dict[line[3]] = line[1]
                            else:  
                                # optional: handle rows with only one value
                                blpr_dict[line[3]] = ""
        else:
            blpr_dict = {}

        
        substring = str(2000) + "-"
        self.record_file = self.prepare_csv()
        # self.clean_csv_edges(self.record_file)
        # print(tqdm(reversed(file.splitlines())))
        # td = input("tdqm")
        if file:
            for line in tqdm(reversed(file.splitlines())):
                try:
                    line = self.stringifyAppend(filename, line, item_dict, disc_dict, dept_dict, type_dict, time_dict, tnsc_dict, paym_dict, blpr_dict)
                    col = line.split(",")
                    #  ? col[8] is date
                    if col[8].strip() != self.date:
                        self.update_monitor_csv()
                    else:
                        if a < 1048575:
                            if(not line == ""):
                                self.csvGenAppend(self.record_file, b, line)
                        else:
                            b += 1
                            if(not line == ""):
                                self.csvGenAppend(self.record_file, b, line)
                            a = 0
                    #print("Finished Processing Line " + str(c) + "!")
                    a += 1
                    c += 1
                except Exception as e:
                    print('Line: %s' % (line))
                    print('Failed to Append. Reason: %s' % ( e))
                    # error = input("There is an error")
                    # hi = input("an error?")

            
        print("Finished converting rd5000")
        # self.clean_csv_edges(self.parentDir + "/record2025.csv")

    def preProc (self, filename):
        normalized = os.path.normpath(filename)

        # Get only the base filename
        base_name = os.path.basename(normalized)   # "a_AYALA-FRN_1_blpr_2025-07-01_20-00_.csv"

        # remove extension if it exists
        name, _ = os.path.splitext(base_name)
        
        try:
            f1 = open(self.parentDir + self.filePaths[0] + "/" + name + ".csv", "r")

            # print(filename)
            # print(f1)
            file = f1.read()
            f1.close()
            # print(file)
            # ih = input("no error?")
            return file
        except FileNotFoundError:
			# If file does not exist, just return empty content
            # print(name)
            # ih = input("error?")
            return ""
        
        # if "blpr" not in filename.lower():
        #     file = file.replace('\n', '')
            # file = file.replace('\r', '')
        # file = file.replace('\x00', '')
        # file = file.replace('/', ' ')
        # file = file.replace('\\~', ' ')
        # file = file.replace('|', ' ')
        # file = file.replace('\\N', '[NULL]')
        # file = file.replace('~', '\n')
        # file = file.replace('\'\'', '"')
    # ? Main function for creating the per line of the master data by appending the necessary fields from the reference files.
    def stringifyAppend(self, filename, line, item, disc, dept, type, time, tnsc, paym, blpr):
        bpcust = False
        x = line.split(",")
        y = []
        # ? 12 is for date.
        columns = [0,2,4,5,6,7,8,11,12,13,18,21,31,32,34,35,36,37]
        for i in range(len(x)):
            temp = ""
            a = False
            if (i in columns):
                if(i == 12):
                    temp = x[i].split(" ")
                    y.append(temp[0])
                    # print(temp)
                    # print(y)
                    # si = input("time?")
                else:
                    if(len(x[i]) >= 2):
                        if(x[i][0] == '0'):
                            a = True
                        if(x[i][1] == '.'):
                            a = False
                    if(len(x[i]) > 2):
                        if(x[i][2] == ':'):
                            a = False
                    if(len(x[i]) > 4):
                        # print(len(x[i]))
                        # print(x[i][4])
                        # print(i)
                        # print(x)
                        # ia = input("why?")
                        if(x[i][4] == '-'):
                            a = False
                    if(i == 11):
                        a = True
                    if(i == 37):
                        a = False
                    if(a == True):
                        temp = '\"=\"\"' + x[i] + '\"\"\"'
                        y.append(temp)
                    else:
                        y.append(x[i])

        # ✅ make sure y has at least 18 elements so y[17] is safe
        while len(y) <= 17:
            y.append("")

        # enforce POS
        y[0] = self.pos
        # pprint.pprint(item)
        # pprint.pprint(y)
        # huh = input("huh")
            
        if (y[2] in item):
            y.append('\"'+item[y[2]]['item_name']+'\"')
            if self.branch in self.new_branches:
                y[7] = '\"'+item[y[2]]['department_code']+'\"'
               
        else:
            y.append("")
        
        # pprint.pprint(item)
        raw = y[7].strip() if len(y) > 7 else ""

        # remove Excel ="..." wrappers and quotes
        code = raw.replace('="', '').replace('"', '').strip()

        if code and code in dept:
            y.append(f"\"{dept[code]}\"")
        else:
            y.append("")
        # pprint.pprint(y)
        # pprint.pprint(item[y[2]])
        # huh = input("huh")
        # disc lookup
        discKey = str(y[10])
        if (discKey in disc):
            y.append('\"'+disc[discKey]+'\"')
        else:
            y.append("")
        
        if (y[11] in type):
            y.append('\"'+type[y[11]]+'\"')
        else:
            y.append("")

        # time lookup
        if y[9][0:1]:
            y.append(time[int(y[9][0:2])])
        else:
            y.append("No Time Record")

        # tnsc + paym + blpr lookup
        if (y[17]):
            #y.append(y[17])
            if (y[17] in tnsc):
                    y.append(tnsc[y[17]])
                    if(tnsc[y[17]] in paym):
                        y.append(paym[tnsc[y[17]]])
                    else:
                        y.append("")
            else:
                y.append("")
                y.append("")
            #print(str(y[17]))
            if (y[17] in blpr):
                y.append(blpr[y[17]])
                bpcust = True
            else:
                y.append("")
        else:
            y.append("")
            y.append("")
            y.append("")
        # print("y17")
        # print(y[17])
        # print("blpr")
        # print(blpr)
        # print("blpry17")
        # print(blpr[y[17]])
        # print("y")
        # print(y)
        # exit();
        # append branch at the end
        y.append(str(self.branch))

        # build final string (clean CSV row)
        string = ",".join(str(item) for item in y)

        return string

    def update_monitor_csv(self, filename="masterData_errorMonitoring.csv"):
        """Append a unique row (pos, branch, date) to an error monitor CSV, creating it if needed."""
        headers = ["pos", "branch", "date"]
        new_entry = [self.pos, self.branch, self.date]

        existing_rows = []

        # Read existing file if it exists
        if os.path.exists(filename):
            with open(filename, "r", newline="") as f:
                reader = csv.reader(f)
                existing_rows = list(reader)

            # Remove header for comparison
            data_rows = existing_rows[1:] if existing_rows else []

            # Check if already logged
            if new_entry in data_rows:
                return  # Already exists → do nothing
        else:
            # File does not exist → create with headers
            with open(filename, "w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(headers)
                writer.writerow(new_entry)
            return

        # Append new unique entry
        with open(filename, "a", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(new_entry)

    def prepare_csv(self):
        """Resolve output CSV path using settings/current_record.txt and ensure it exists with headers.
        
        Notes:
        - When Rewrite mode is active, GUI stores a versioned path in settings/current_record.txt.
        - When Append mode is active, the pointer is reset to record2025.csv.
        - This function reads that pointer and writes headers if the target file is missing/empty.
        """
        processed_dir = self.parentDir + "/"
        default_record = os.path.join(processed_dir, "record2025.csv")
        header_file = os.path.join(processed_dir, "aaa_headers.csv")
        pointer_path = os.path.join(self.parentDir, "settings", "current_record.txt")
        record_file = default_record
        try:
            # Use pointer if present to direct output to the selected record file
            if os.path.exists(pointer_path):
                with open(pointer_path, "r", encoding="utf-8") as f:
                    p = f.read().strip()
                    if p:
                        record_file = p
        except Exception:
            record_file = default_record
        try:
            # Ensure destination directory exists
            os.makedirs(os.path.dirname(record_file), exist_ok=True)
        except Exception:
            pass
        # Create file and write header when missing/empty
        if (not os.path.exists(record_file)) or (os.path.getsize(record_file) == 0):
            with open(record_file, "w", encoding="utf-8") as f_out:
                if os.path.exists(header_file):
                    with open(header_file, "r", encoding="utf-8") as f_header:
                        f_out.write(f_header.read().strip() + "\n")
        return record_file
    def csvGenAppend(self, filename, part, line):
        """Append a processed line to the open record file."""
        with open(self.record_file, "a", encoding="utf-8") as f_out:
            f_out.write(str(line) + "\n")
    def clean_csv_edges(self, file_path):
        """Trim empty first/last lines in a CSV file and rewrite cleaned content."""
        with open(file_path, "r", encoding="utf-8") as f:
            lines = [line.rstrip("\n") for line in f]

        # Remove first row if empty
        if lines and lines[0].strip() == "":
            lines.pop(0)

        # Remove last row if empty
        if lines and lines[-1].strip() == "":
            lines.pop()

        # Rewrite file without extra empty rows
        with open(file_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + ("\n" if lines else ""))
    # def append(self):
    #     if os.path.isfile("./append.csv"):
    #         date_start = '2025-01-01'
    #         date_end = '2027-12-31'
    #         prev_table = CSVProcessor("record2025.csv")
    #         prev_table = prev_table.filter('','','',date_start,date_end)
    #         prev_table.to_csv('record_2025_temp.csv',index = False)
    
    #         print("Append")
    #         new = CSVProcessor("append.csv")
    #         new_table = new.filter('','','','2025-05-01','2025-05-31')
    #         if not new_table.empty:
    #             vertical_concat = pd.concat([prev_table, new_table], axis=0)
    #             vertical_concat.to_csv('record_2025_new.csv',index = False)
    #         #os.remove("./append.csv")
    #     else:
    #         print("No Append")
