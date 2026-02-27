from tkinter import messagebox, filedialog
import pandas as pd
import os
from stat import *
from hurry.filesize import size as sz
import datetime
import re

##PROCESSES OUTPUT OF SEEKLIGHT AI AND CONVERTS TO WORKABLE MARC IN XLSX FOR IMPORT TO ALMA##

#Hard code list of headings to remove from records
XHEADINGS = ["Canada","Twenty-second century","Twenty-first century","Twentieth century","Nineteenth century","Eighteenth century","Seventeenth century","Sixteenth century","Fifteenth century","Government regulation of ...","Government publications","Canadian provinces","Commons","Nineteen eighties","Nineteen fifties","Nineteen forties","Nineteen hundreds (Decade)","Nineteen nineties","Nineteen seventies","Nineteen sixties","Nineteen tens","Nineteen thirties","Nineteen twenties","Two thousand, A.D."]

def checkfiles(pdfs,df):
    #Create list of filenames to check against PDFs in directory
    filenames = df["Filename"].to_list()
    print(pdfs, filenames)
    #Check for PDF filenames in Seeklight export
    extrapdfs = []
    for pdf in pdfs:
        if pdf not in filenames:
            extrapdfs.append(pdf)
    #Check Seeklight export filenames against PDFs
    extrarecords = []
    for file in filenames:
        if file not in pdfs:
            extrarecords.append(file)
    if len(extrapdfs) > 0:
        messagebox.showwarning(title="Error",message=f'Error: Missing records for the following PDFs: {extrapdfs}.\nPlease verify contents of directory and Seeklight export and try again.')
        exit()
    if len(extrarecords) > 0:
        messagebox.showwarning(title="Error",message=f'Error: Missing PDFs from directory for the following records: {extrarecords}.\nPlease verify contents of directory and Seeklight export and try again.')
        exit()
    if len(extrarecords) == 0 and len(extrapdfs) == 0:
        print("Contents of directory confirmed. Processing...")

def filemerge(files):
    ##Should default to UTF-8, but may need to change this to with open(file, encoding="utf-8) as fh: df = pd.read_excel(fh)
    f1 = pd.read_excel(files[0], names=["SSID","Filename","File Count","Title","Creator","Volume","Issue","Date","Language","Description","Subject","Type","Coverage","Format","Publisher","Medium","Technique","Material","Measurements","Style","Culture","Period","Location","Named Entities","Keywords","Resource Type","Media URL"])
    for file in files[1:]:
        #Error checking for filetype
        try:
            pd.read_excel(file)
            print("File read successfully. Continuing with processing...")
        except Exception as e:
            messagebox.showwarning(title="Error",message=f'Error: File read of {file} unsuccessful due to the following: {e}\nPlease ensure all Excel files in directory are valid Seeklight exports in .xlsx format.')
            exit()
        ##Should default to UTF-8, but may need to change this to with open(file, encoding="utf-8) as fh: df = pd.read_excel(fh)
        f2 = pd.read_excel(file, names=["SSID","Filename","File Count","Title","Creator","Volume","Issue","Date","Language","Description","Subject","Type","Coverage","Format","Publisher","Medium","Technique","Material","Measurements","Style","Culture","Period","Location","Named Entities","Keywords","Resource Type","Media URL"])
        f1 = pd.concat([f1, f2], ignore_index=True, sort=False)
    return (f1)

def build008(row):
    ##Add date from current, formatted to 6 digits
    currentdate = datetime.date.today().strftime("%y%m%d")
    field008 = currentdate + "s"
    #add date
    field008 = field008 + str(row["Date"])[:4]
    field008 = field008 + "####"
    #add place of publication (Canadian provinces and territories, default to Canada if no province listed)
    # if "british columbia" in str(row["Location"]).lower() or "b.c." in str(row["Location"]).lower():
    #     field008 = field008 + "bcc"
    # elif "alberta" in str(row["Location"]).lower():
    #     field008 = field008 + "abc"
    # elif "saskatchewan" in str(row["Location"]).lower():
    #     field008 = field008 + "snc"
    # elif "manitoba" in str(row["Location"]).lower():
    #     field008 = field008 + "mbc"
    # elif "ontario" in str(row["Location"]).lower():
    #     field008 = field008 + "onc"
    # elif "quebec" in str(row["Location"]).lower() or "québec" in str(row["Location"]).lower():
    #     field008 = field008 + "quc"
    # elif "new brunswick" in str(row["Location"]).lower():
    #     field008 = field008 + "nkc"
    # elif "prince edward island" in str(row["Location"]).lower() or "pei" in str(row["Location"]).lower():
    #     field008 = field008 + "pic"
    # elif "nova scotia" in str(row["Location"]).lower():
    #     field008 = field008 + "nsc"
    # elif "newfoundland" in str(row["Location"]).lower() or "labrador" in str(row["Location"]).lower():
    #     field008 = field008 + "nfc"
    # elif "yukon" in str(row["Location"]).lower():
    #     field008 = field008 + "ykc"
    # elif "northwest territories" in str(row["Location"]).lower() or "nwt" in str(row["Location"]).lower():
    #     field008 = field008 + "ntc"
    # elif "nunavut" in str(row["Location"]).lower():
    #     field008 = field008 + "nuc"
    # else:
    #     field008 = field008 + "-cn"
    ##Hard coding location
    field008 = field008 + "onc"
    field008 = field008 + "####|o####f00|#0#"
    #add language based on content of Language field
    # if "|" in str(row["Language"]):
    #     field008 = field008 + "mul"
    # elif "french" in str(row["Language"]).lower():
    #     field008 = field008 + "fre"
    # elif "english" in str(row["Language"]).lower():
    #     field008 = field008 + "eng"
    # else:
    #     field008 = field008 + "und"
    ##Language set to English only
    field008 = field008 + "eng"
    field008 = field008 + "#d"
    return field008

def getfilesize(row):
    #Error checking to ensure all PDFs present
    try:
        filesize = os.stat(str(row["Filename"]).split("/")[-1]).st_size
    except Exception as e:
        messagebox.showinfo(title="Error",message="Error: Unable to proceed. Please ensure all PDFs are present and filenames match Seeklight output.")
        exit()
    #Add B to suffix; formatting will be fixed on import to Alma
    filesize = sz(filesize) + "B"
    return(filesize)

def removeheadings(subjects):
    #Split headings into list
    headings = subjects.split("|")
    #Create empty list to receive indices of terms to delete
    delindex = []
    #For each heading in the list, check against list of headings to remove
    for heading in headings:
        print("heading is ",heading)
        for xheading in XHEADINGS:
            if heading.strip() == xheading:
                delindex.append(headings.index(heading))
                continue
    #Reverse index to work backwards through list (preserves indices while removing terms)
    delindex.sort(reverse=True)
    for i in delindex:
        headings.pop(i)
    #Return headings as string (will be split on import to Alma)
    return "%".join(headings)

def mdprocess(i, row, MARCdf):
    #set fixed/default fields (LDR, 040, 300, 336, 337, 338, 347 (partial), 533 (partial), 588 (partial), 901)
    MARCdf["LDR"][i] = "#####nam#a22#####7i#4500"
    MARCdf["006"][i] = "m#####o##d#f######"
    MARCdf["007"][i] = "cr#bn#|||a||||"
    MARCdf["040$a"][i] = "CaOOP"
    MARCdf["040$b"][i] = "eng"
    MARCdf["040$e"][i] = "rda"
    MARCdf["040$c"][i] = "CaOOP"
    MARCdf["055 3$a"][i] = "J103 H7"
    MARCdf["24510$a"][i] = "Sessional paper"
    MARCdf["300$a"][i] = "1 online resource"
    MARCdf["336$a"][i] = "text"
    MARCdf["336$b"][i] = "txt"
    MARCdf["336$2"][i] = "rdacontent"
    MARCdf["337$a"][i] = "computer"
    MARCdf["337$b"][i] = "c"
    MARCdf["337$2"][i] = "rdamedia"
    MARCdf["338$a"][i] = "online resource"
    MARCdf["338$b"][i] = "cr"
    MARCdf["338$2"][i] = "rdacarrier"
    MARCdf["347$a"][i] = "text file"
    MARCdf["347$2"][i] = "rdaft"
    MARCdf["3479 $b"][i] = "PDF"
    MARCdf["533$b"][i] = "Ottawa :"
    MARCdf["533$c"][i] = "Library of Parliament,"
    MARCdf["533$5"][i] = "CaOOP"
    MARCdf["588$5"][i] = "CaOOP"
    #MARCdf["7109 $a"][i] = "Canada."
    #MARCdf["7109 $b"][i] = "Parliament.%House of Commons.%Office of the Government House Leader,"
    ##Alma does not seem to accept 710$e when loading? Not sure why, but can add with normalization rule for the time being.
    #MARCdf["7109 $e"][i] = "issuing body."
    MARCdf["901$a"][i] = "SESSIONPAP"
    #construct 008
    field008 = build008(row)
    MARCdf["008"][i] = field008
    #set variable fields (035, 041, 100, 245, 264, 347 (partial), 490, 520, 533$d, 600, 650, 700, 830, 988)
    MARCdf["035$z"][i] = "(JSTOR)" + str(row["SSID"])
    # if "|" in str(row["Language"]):
    #     MARCdf["041$a"][i] = str(row["Language"])
    ##Hard code language to eng and fre
    MARCdf["0410 $a"][i] = "eng|fre"
    ##Add sessional paper number to call number
    MARCdf["055 3$b"][i] = str(str(row["Filename"]).split("/")[-1]).rstrip(".pdfa")
    #split Creator into 110 and 710 at | if present
    # if "|" in row["Creator"]:
    #     MARCdf["1102 $a"][i] = str(row["Creator"])[:str(row["Creator"]).index('|')]
    #     MARCdf["710$a"][i] = str(row["Creator"])[str(row["Creator"]).index('|'):]
    # else:
    #     MARCdf["1102 $a"][i] = str(row["Creator"])
    ##Putting all creators/named entities in 700 by default
    MARCdf["700$a"][i] = str(row["Creator"]) + "%" + str(re.sub(r'\|','%',str(row["Named Entities"])))
    #MARCdf["24510$a"][i] = str(row["Title"])
    ##Replace title with filename
    MARCdf["24510$b"][i] = str(str(row["Filename"]).split("/")[-1]).rstrip(".pdfa")
    # MARCdf["264 1$a"][i] = str(row["Location"])
    # MARCdf["264 1$b"][i] = str(row["Publisher"])
    ##Hard-coding publisher
    MARCdf["264 1$a"][i] = "[Ottawa]:"
    MARCdf["264 1$b"][i] = "[House of Commons],"
    MARCdf["264 1$c"][i] = str(row["Date"])[:4] + "."
    MARCdf["3479 $c"][i] = str(getfilesize(row))
    MARCdf["4901 $a"][i] = "Sessional paper / House of Commons = Document parlementaire / Chambre des communes ; "
    MARCdf["4901 $v"][i] = str(str(row["Filename"]).split("/")[-1]).rstrip(".pdfa")
    ##533$d using datetime to fetch current year
    currentyear = datetime.date.today().strftime("%Y")
    MARCdf["533$d"][i] = str(currentyear)
    #MARCdf["520$a"][i] = str(row["Description"])
    ##520 removed at request
    #MARCdf["610$a"][i] = str(row["Named Entities"])
    ##All named entities to 700 with creator
    #Remove unwanted headings
    headings = removeheadings(str(row["Subject"]))
    MARCdf["650 0$a"][i] = headings
    ##Removing all headings --Put back in on 20260219 -KB
    #Conditional handling of 533$a/n, 588$a, 830$a, defaulting to English unless French specified in 040$b (currently hard-coded to eng, so never fre)
    if MARCdf["040$b"][i] == "fre":
        MARCdf["533$a"][i] = "Reproduction électronique."
        MARCdf["533$n"][i] = "Reproduction électronique de documents imprimés détenus par la Bibliothèque du Parlement."
        MARCdf["588$a"][i] = "Certaines métadonnées de cette notice bibliographique ont été générées à l’aide de l’IA."
        MARCdf["830 0$a"][i] = "Document parlementaire (Canada. Parlement. Chambre des communes) ; "
    else:
        MARCdf["533$a"][i] = "Electronic reproduction."
        MARCdf["533$n"][i] = "Electronic reproduction from printed material held by the Library of Parliament."
        MARCdf["588$a"][i] = "Portions of the metadata in this bibliographic record were created with the help of AI."
        MARCdf["830 0$a"][i] = "Sessional paper (Canada. Parliament. House of Commons) ; "
    MARCdf["830 0$v"][i] = str(str(row["Filename"]).split("/")[-1]).rstrip(".pdfa")
    MARCdf["988$a"][i] = str(row["Filename"]).split("/")[-1]
    #ignore empty fields for manual cleanup (110, 610, 710)
    return MARCdf

def main():
    #prompt user to select directory for processing
    messagebox.showinfo(title="Seeklight Processing",message="Please select a directory to process. Ensure all Seeklight files (Excel and PDF) are in the same folder.")
    direc = filedialog.askdirectory(mustexist=True)
    #Change active directory to selected
    os.chdir(direc)
    #List all files in directory
    allfiles = os.listdir(direc)
    xlfiles = []
    pdfs = []
    #Extract all Excel files and PDFs for handling
    for file in allfiles:
        if ".xlsx" in file:
            xlfiles.append(file)
        elif ".pdf" in file or ".pdfa" in file:
            pdfs.append(file)
        else:
            strays = messagebox.askokcancel(title="Error",message=f'Potential Error: "{file}" is not of an allowed type. Please confirm that folder contains only Seeklight export(s) in .xlsx and PDFs (or subdirectories meant to be ignored). Click OK to continue, Cancel to abort.')
            if strays == False:
                exit()
    #check file(s) and merge into single file for processing if more than one; quit if no file selected.
    if len(xlfiles) > 1:
        df = filemerge(xlfiles)
        print(df)
    elif len(xlfiles) == 1:
        file = xlfiles[0]
        ##Should default to UTF-8, but may need to change this to with open(file, encoding="utf-8) as fh: df = pd.read_excel(fh)
        df = pd.read_excel(file, names=["SSID","Filename","File Count","Title","Creator","Volume","Issue","Date","Language","Description","Subject","Type","Coverage","Format","Publisher","Medium","Technique","Material","Measurements","Style","Culture","Period","Location","Named Entities","Keywords","Resource Type","Media URL"])
        print(df)
    else:
        print("Error: Please ensure at least one valid Seeklight export (Excel file) is present.")
        quit
    #Check PDFs against filenames in Seeklight export
    checkfiles(pdfs, df)
    #create dataframe to receive reformatted data
    reci = len(df)
    ##"520$a" removed at request
    ##710 currently only used for preset data (ind1=9 used to isolate and merge on import); if using in future, add another set using ind1=2.
    MARCdf = pd.DataFrame(columns=["LDR","006","007","008","035$z","040$a","040$b","040$e","040$c","0410 $a","055 3$a","055 3$b","1001 $a","1102 $a","24510$a","24510$b","264 1$a","264 1$b","264 1$c","300$a","336$a","336$b","336$2","337$a","337$b","337$2","338$a","338$b","338$2","347$a","347$2","3479 $b","3479 $c","4901 $a","4901 $v","533$a","533$b","533$c","533$d","533$n","533$5","588$a","588$5","600$a","610$a","650 0$a","700$a","7109 $a","7109 $b","830 0$a","830 0$v","901$a","988$a"],index=range(reci))
    print("Processing dataframe...")
    #iterate through dataframe, populating MARC fields
    i = 0
    for index, row in df.iterrows():
        print(index, row)
        MARCdf = mdprocess(i, row, MARCdf)
        i += 1
    #prompt user to save output
    messagebox.showinfo(title="Seeklight Processing",message="Please select a location and name for the output.")
    output = filedialog.asksaveasfilename()
    #write MARCdf to new Excel sheet
    MARCdf.to_excel(f'{output}.xlsx', index=False)
    messagebox.showinfo(title="",message=f'Output saved to {output}.xlsx. Process complete.')

if __name__ == "__main__":

    main()
