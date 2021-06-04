import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import ttk, filedialog as fd


class orphanGUI(object):
    def __init__(self):
        self.win = tk.Tk()
        self.win.geometry("700x500")
        self.win.resizable(1000, 1500)
        self.win.title("Kroger Orphan Template Tool")
        self.orphan_gui_widgets()

    def file_opener(self, filetype):
        input = fd.askopenfilename(initialdir="/")
        if filetype == 'd':
            dispFile.set(input)
            #print(dispFile)
        if filetype == 'm':
            unmatchFile.set(input)
            #print(unmatchFile)
        if filetype == 'h':
            hierFile.set(input)
            # print(unmatchFile)

    def save_location(self):
        input = fd.askdirectory(initialdir="/")
        fileDest.set(input+'/')

    def orphan_gui_widgets(self):
        global dispFile
        dispFile = tk.StringVar()
        global unmatchFile
        unmatchFile = tk.StringVar()
        global hierFile
        hierFile = tk.StringVar()
        global fileDest
        fileDest = tk.StringVar()
        tabcontrol = ttk.Notebook(self.win)
        tab1 = ttk.Frame(tabcontrol)
        tabcontrol.add(tab1, text='Kroger Orphans')
        tabcontrol.pack(expand=1, fill="both")

        self.kOrphan = ttk.LabelFrame(tab1, text='Specify the Input File Locations')
        self.kOrphan.grid(column=0, row=0, padx=8, pady=4)

        ttk.Label(self.kOrphan, text="Disposition Report").grid(column=0, row=0, sticky="W")
        ttk.Label(self.kOrphan, text="Unmatched Report").grid(column=0, row=1, sticky="W")
        ttk.Label(self.kOrphan, text="Hierarchy Report").grid(column=0, row=2, sticky="W")
        ttk.Label(self.kOrphan, text="File Destination").grid(column=0, row=3, sticky="W")

        self.getdispbtn = ttk.Button(self.kOrphan, text="Browse", command=lambda: self.file_opener('d'))
        self.getdispbtn.grid(column=3, row=0)

        self.dispReportFile = ttk.Entry(self.kOrphan, width=50, textvariable=dispFile)
        self.dispReportFile.grid(column=6, row=0, sticky="WE")

        self.getunmatchbtn = ttk.Button(self.kOrphan, text="Browse", command=lambda: self.file_opener('m'))
        self.getunmatchbtn.grid(column=3, row=1)

        self.unmatchedReportFile = ttk.Entry(self.kOrphan, width=50, textvariable=unmatchFile)
        self.unmatchedReportFile.grid(column=6, row=1, sticky="WE")

        self.gethierbtn = ttk.Button(self.kOrphan, text="Browse", command=lambda: self.file_opener('h'))
        self.gethierbtn.grid(column=3, row=2)

        self.hierReportFile = ttk.Entry(self.kOrphan, width=50, textvariable=hierFile)
        self.hierReportFile.grid(column=6, row=2, sticky="WE")

        self.getfdestbtn = ttk.Button(self.kOrphan, text="Browse", command=lambda: self.save_location())
        self.getfdestbtn.grid(column=3, row=3)

        self.fileDestLoc = ttk.Entry(self.kOrphan, width=50, textvariable=fileDest)
        self.fileDestLoc.grid(column=6, row=3, sticky="WE")

        self.buildbtn = ttk.Button(self.kOrphan, text="Build Template", command=lambda: self.template_build(dispFile, unmatchFile, hierFile, fileDest))
        self.buildbtn.grid(column=0, row=4)

    # This function will build the spreadsheets for batch uploads I also initiates Order 66 from Emperor Palpatine
    def template_build(self, dispFile, unmatchFile, hierFile, fileDest):
        # creating variables to hold the addresses of the files used to build the templates
        hierarchy_sheet = hierFile.get()
        unmatched_sheet = unmatchFile.get()
        disp_sheet = dispFile.get()
        regulated_sheet = dispFile.get()
        f_dest = fileDest.get()

        # defining the columns that are used in the process so that those ar the only ones imported
        dispcols = ["Chain of Custody", "Donor First Name", "Donor Last Name", "SSN", "Collection Date",
                    "Site State", "Test Type", "Test Reason", "Panel ID", "Panel Description", "Master", "Sub"]

        # creating dictionaries of what the simplified vectors and reasons are to later replace with user readable
        testVector = {"Urine": "Urine Drug Test",
                      "Oral Fluid": "Oral Fluid Test",
                      "eCup": "Instant Urine Test",
                      "Breath": "Instant Oral Fluid Test"}
        testReason = {"PREEMP": "Pre-employment",
                      "FOLLOWUP": "Follow-up",
                      "OTHER": "Others",
                      "POSTACCIDENT": "Post Accident",
                      "RANDOM": "Random",
                      "RETURNTODUTY": "Return to Duty",
                      "REASONABLESUSPICION": "Reasonable Suspicion"}
        # Creating variables to hold the data from the sheets
        KrogHier = pd.read_excel(hierarchy_sheet, converters={'Division': lambda x: str(x)})
        KrogUnmatch = pd.read_excel(unmatched_sheet, sheet_name='Sheet2')
        KrogDisp = pd.read_excel(disp_sheet, usecols=dispcols)
        KrogReg = pd.read_excel(regulated_sheet, usecols=["Chain of Custody", "Regulated"])

        # Renaming the some columns to be able to be joined on
        KrogDisp.rename(columns={"Chain of Custody": "Specimen ID"}, inplace=True)
        KrogReg.rename(columns={"Chain of Custody": "Specimen ID"}, inplace=True)

        # Combining the Master and Sub together to create one concatenated value and then removing the old columns
        KrogDisp["MasterSub"] = KrogDisp["Master"].astype(str) + "-" + KrogDisp["Sub"].astype(str)
        KrogDisp["Collection Date"] = KrogDisp["Collection Date"].dt.strftime('%m/%d/%Y')
        del KrogDisp["Master"]
        del KrogDisp["Sub"]

        # Creating a data variable to add to the output file names as well as a value for Nonexistent names
        currDate = pd.to_datetime("today").strftime('%m/%d/%Y')
        flNameReplace = "NOT FOUND"

        # Checks input string for numeric characters
        def has_numbers(inputString):
            return any(char.isdigit() for char in inputString)

        # Removes non-numeric characters from string
        def remove_char(inputString):
            if not inputString.isnumeric():
                numeric_filter = filter(str.isdigit, inputString)
                numeric_string = "".join(numeric_filter)
                inputString = numeric_string
            return inputString

        # Removes non-numeric characters from SSN. This is done to prevent other federal IDs from being input in for SSN
        KrogDisp["SSN"] = KrogDisp["SSN"].apply(lambda x: remove_char(str(x)))

        # Joining the data frames
        krogJoin = pd.merge(KrogUnmatch, KrogDisp, on="Specimen ID")
        mergeData = pd.merge(krogJoin, KrogHier, on="MasterSub", how="outer")
        mergeData.dropna(subset=['Specimen ID'], inplace=True)
        mergeData = pd.merge(mergeData, KrogReg, on="Specimen ID", how="inner")

        # Cleaning and converting data in the data frames, as well as constructing the ne data frame
        mergeData.insert(0, "Account Code", (mergeData["Test Reason"].apply(lambda x: "KRGFM" if x == "PREEMP" else "KRGFM01")))
        mergeData.insert(1, "User Login", (mergeData["Account Code"].apply(lambda x: "dutUx390" if x == "KRGFM" else "duTeA608")))
        mergeData.insert(2, "Optional: User Group Reference ID", mergeData["Division"])
        del mergeData["Division"]
        mergeData.insert(3, "Group Code", "")
        mergeData.insert(4, "Group Name", "")
        mergeData.insert(5, "First Name", (mergeData['Donor First Name'].apply(lambda x: flNameReplace if any(map(str.isdigit, str(x))) == True else x)))
        del mergeData['Donor First Name']
        mergeData.insert(6, "Middle Name", "")
        mergeData.insert(7, "Last Name", (mergeData['Donor Last Name'].apply(lambda x: flNameReplace if has_numbers(str(x)) == True else x)))
        del mergeData['Donor Last Name']
        mergeData.rename(columns={"SSN": "DispSSN"}, inplace=True)
        mergeData.insert(8, "SSN", (mergeData["DispSSN"].apply(lambda x: "" if (x.isnumeric() == False or len(x) < 9) else x)))
        del mergeData['DispSSN']
        mergeData.insert(9, "Applicant/Employee does not have a SSN", (mergeData["SSN"].apply(lambda x: "" if len(x) == 9 else "Yes")))
        mergeData.insert(10, "Date Of Birth", "01/01/1900")
        mergeData.insert(11, "Country", "US")
        mergeData.rename(columns={"State": "DispState"}, inplace=True)
        mergeData.insert(12, "State", (mergeData["Brand"].apply(lambda x: "TX" if x == "Southwest" else "SC")))
        del mergeData["DispState"]
        mergeData.insert(13, "Street Address", "1 Main St")
        del mergeData["City"]
        mergeData.insert(14, "City", (mergeData["State"].apply(lambda x: "Chapin" if x == "SC" else "Austin")))
        mergeData.insert(15, "ZIP Code", (mergeData["State"].apply(lambda x: "29036" if x == "SC" else "78701")))
        mergeData.insert(16, "Email Address", "")
        mergeData.insert(17, "Phone", "9999999999")
        del mergeData["Reason for Test"]
        mergeData.insert(18, "Reason for test", mergeData["Test Reason"])
        mergeData["Reason for test"].replace(testReason, inplace=True)
        mergeData.insert(19, "Registration expires in ?", "")
        mergeData.insert(20, "Test Panel", "5 Panel")
        mergeData.insert(21, "Sample Type", mergeData["Test Type"])
        mergeData["Sample Type"].replace(testVector, inplace=True)
        mergeData.insert(22, "Coordination Type", "None")
        mergeData.rename(columns={"Specimen ID": "DispSpecimen ID"}, inplace=True)
        mergeData.insert(23, "Specimen ID", mergeData["DispSpecimen ID"])
        mergeData.insert(24, "Job Location: Nickname", "")
        mergeData.insert(25, "Job Location: Country", "")
        mergeData.insert(26, "Job Location: City", "")
        mergeData.insert(27, "Job Location: State/Province", "")
        mergeData.insert(28, "Job Location: Zip/Postal Code", "")
        mergeData.insert(29, "Job Location: Confirm Zip Code is a Physical Address", "")
        mergeData.insert(30, "Job Location: Region", "")
        mergeData.insert(31, "Flex Field: Division", "")
        mergeData.insert(32, "Flex Field: Location", "Orphan")
        mergeData.insert(33, "Flex Field: Requester Email", "Orphan")
        mergeData.insert(34, "Flex Field: Exempt/Non Exempt", "Exempt")
        mergeData.insert(35, "Flex Field: Job Title", "Orphan")
        mergeData.insert(36, "Flex Field: Candidate ID", "Orphan")
        mergeData.insert(37, "Flex Field: App ID", "")
        mergeData.insert(38, "Flex Field: NotificationEmail", "")

        # Filling in an missing user group data with a fixed value
        mergeData['Optional: User Group Reference ID'] = mergeData['Optional: User Group Reference ID'].fillna("021")

        # Removing unneeded columns
        del mergeData["DHS Client Name"]
        del mergeData["Employer"]
        del mergeData["Applicant Name"]
        del mergeData["Applicant Last 4 SSN"]
        del mergeData["Spec ID"]
        del mergeData["Service"]
        del mergeData["Sub Request Code"]
        del mergeData["Submitted Date"]
        del mergeData['Test Reason']
        del mergeData["Panel ID"]
        del mergeData["DispSpecimen ID"]
        del mergeData["Collection Date"]
        del mergeData["Site State"]
        del mergeData["Test Type"]
        del mergeData["Panel Description"]
        del mergeData["MasterSub"]
        del mergeData["LocationName"]
        del mergeData["Corporate"]
        del mergeData["Brand"]

        # Adding the first row of the template
        mergeData.loc[-1] = [
            "Required: Use the same value that is used in the Company ID field when logging into HireRight",
            "Required: Use the same value that is used in the username field when logging into HireRight",
            "Optional: User Group Reference ID",
            "Optional: User Group Code",
            "Optional: User Group Name",
            "Required: Applicant's/Employee's first name",
            "Optional: Applicant's/Employee's middle name",
            "Required: Applicant's/Employee's last name",
            "May be Required: Unless applicant does not have a SSN, Applicant's/Employee's Social Security Number. Enter 9 digits",
            "May be Required: If SSN not entered, If the Applicant/Employee has not been issued a SSN enter Yes",
            "Required: Applicant's/Employee's DOB, MM/DD/YYYY",
            "Required: Two letter country code. Use US for USA, CA for Canada, GB for United Kingdom. Refer to ISO3166-1 for all countries",
            "Required: Applicant's/Employee's Two letter US state or a complete state or province name",
            "Required: Applicant's/Employee Street Address",
            "Required: Applicant's/Employee City",
            "Required: Applicant's/Employee's Zip Code",
            "Optional: Applicant's/Employee's E-mail",
            "Required: Applicant's/Employee's phone number",
            "Required: Select Reason from the Dropdown",
            "Optional: May enter if using eCOC, If left blank will use default, Enter numeral from 1 to 60",
            "Required: Select the Test Panel from the Dropdown",
            "Required: Select the Sample Type from the Dropdown",
            "Required: Select the Coordination Type from the Dropdown.",
            "Optional: If you know the Specimen Id number for the drug test please enter it.",
            "Optional: If left Blank will use customer default: This is the name associated with the Job locations address (for example Riverside Facility or Eastside Plant)",
            "May be Required: Required unless location was preconfigured or using default location, Two letter country code. Use US for USA, CA for Canada, GB for United Kingdom. Refer to ISO3166-1 for all countries:Country in which the candidate will work",
            "May be Required: Required unless location was preconfigured or using default location, City where the candidate will work",
            "May be Required: Required unless location was preconfigured or using default location, State where the candidate will work.Two letter US state or a complete state or province name",
            "May be Required: Required unless location was preconfigured or using default location, Zip code where the candidate will work",
            "May be Required: If the Zip code provided was rejected because it belongs to a PO Box, enter Yes below if you can confirm it is actually a Zip code for a Physical Location",
            "May be Required: When entering new location if Country is UK, Region is required, England, Northern Ireland, Scotland or Wales",
            "Optional",
            "Required",
            "Required",
            "Required",
            "Required",
            "Required",
            "Optional",
            "Optional",
            ""
        ]
        mergeData.index = mergeData.index + 1
        mergeData = mergeData.sort_index()

        # Replacing missing first name with NaN, ssn that are 0 with empty string, and dropping the rows with missing firstnames
        mergeData['First Name'].replace('', np.nan, inplace=True)
        mergeData['SSN'].replace("0", "", inplace=True)
        mergeData.dropna(subset=['First Name'], inplace=True)

        # Creating data frames for the separate Kroger accounts
        KRGFM = mergeData[mergeData["Account Code"] == "KRGFM"]
        KRGFM01 = mergeData[mergeData["Account Code"] == "KRGFM01"]

        # Splitting the data frames into new data frames that are parsed by account and DOT/nonDOT
        KRGFMDOT = KRGFM[KRGFM["Regulated"] == "DOT"]
        KRGFMDOT.loc[0] = mergeData.loc[0]
        KRGFMDOT = KRGFMDOT.sort_index()

        KRGFMND = KRGFM[KRGFM["Regulated"] == "Non-DOT"]
        KRGFMND.loc[0] = mergeData.loc[0]
        KRGFMND = KRGFMND.sort_index()

        KRGFM01DOT = KRGFM01[KRGFM01["Regulated"] == "DOT"]
        KRGFM01DOT.loc[0] = mergeData.loc[0]
        KRGFM01DOT = KRGFM01DOT.sort_index()

        KRGFM01ND = KRGFM01[KRGFM01["Regulated"] == "Non-DOT"]
        KRGFM01ND.loc[0] = mergeData.loc[0]
        KRGFM01ND = KRGFM01ND.sort_index()

        # Dropping the Regulated column from the parsed data frames
        del KRGFMDOT["Regulated"]
        del KRGFMND["Regulated"]
        del KRGFM01DOT["Regulated"]
        del KRGFM01ND["Regulated"]

        # Creating files for the batches
        KRGFMDOT.to_excel(f_dest+"KRGFM DOT Orphan Batch {}.xlsx".format(currDate.replace("/", "-", 3)), index=False)
        KRGFMND.to_excel(f_dest+"KRGFM Non-DOT Orphan Batch {}.xlsx".format(currDate.replace("/", "-", 3)), index=False)
        KRGFM01DOT.to_excel(f_dest+"KRGFM01 DOT Orphan Batch {}.xlsx".format(currDate.replace("/", "-", 3)), index=False)
        KRGFM01ND.to_excel(f_dest+"KRGFM01 Non-DOT Orphan Batch {}.xlsx".format(currDate.replace("/", "-", 3)), index=False)
        mergeData.to_excel(f_dest+" Kroger Master Orphan Batch {}.xlsx".format(currDate.replace("/", "-", 3)), index=False)


        print("End Function")

if __name__ == '__main__':
    print('works')

orphanGUI = orphanGUI()
orphanGUI.win.mainloop()
