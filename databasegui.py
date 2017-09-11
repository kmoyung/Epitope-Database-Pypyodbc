import tkinter
from tkinter import *
from tkinter import filedialog
import pypyodbc

# File Dialog
diag = tkinter.Tk()
diag.withdraw()
diag.filename = filedialog.askopenfilename(initialdir=r"C:\Users\kmoyung\SharePoint\IMBD - Documents\Epitope Prediction\
Prediction Databases"
                                          , title="Select a database")
diag.filename = diag.filename.replace("/","\\")


# GUI
top = tkinter.Tk()
top.title("Database Access")
top.geometry("800x350")

disp_location = Label(top, text = "Database is located in: " + diag.filename)
disp_location.pack()

seqon = IntVar(top)
alleleon = IntVar(top)
antigenon = IntVar(top)
affinityon = IntVar(top)
lengthon = IntVar(top)
poson = IntVar(top)

header = []

# Access the database and run query
def access():

    if entryallele.get() is "":
        finish.config(text = "Don't leave any fields blank!", fg = 'red')
        raise Exception("Please fill all empty fields!")

    if entryantigen.get() is "":
        finish.config(text = "Don't leave any fields blank!", fg = 'red')
        raise Exception("Please fill all empty fields!")


    querycode = 'SELECT '

    print(querycode)

    if seqon.get() == 1:
        querycode = querycode + "t1.Sequence, "
        header.append("Sequence")

    if alleleon.get() == 1:
        querycode = querycode + "t1.AlleleName, "
        header.append("AlleleName")

    if antigenon.get() == 1:
        querycode = querycode + "t1.AntigenName, "
        header.append("AntigenName")

    if affinityon.get() == 1:
        querycode = querycode + "t1.nM, "
        header.append("nM")

    if lengthon.get() == 1:
        querycode = querycode + "t1.SequenceLength, "
        header.append("SequenceLength")

    if poson.get() == 1:
        querycode = querycode + "t1.StartingPosition, t1.EndingPosition, "
        header.append("StartingPosition")
        header.append("EndingPosition")


    # Remove extra comma
    querycode = querycode[:-2]

    querycode = querycode + " FROM [Epitope Database Master Merged] as t1"

    if entryallele.get() is not "":
        querycode = querycode + " WHERE t1.AlleleName = '" + entryallele.get() + "'"

    if entryantigen.get() is not "":
        querycode = querycode + " AND t1.AntigenName = '" + entryantigen.get() + "'"

    print(querycode)

    # Select query save location
    # top.savefolder = filedialog.askdirectory(initialdir=r"C:\Users\kmoyung\Desktop", title="Select save location")
    # top.savefolder = top.savefolder.replace("/","\\")
    # top.savefolder = top.savefolder + "\guitest.csv"

    top.savefile = filedialog.asksaveasfilename(initialdir="/", title = "Name the file and save to a location", filetypes = (("CSV Files", "*.csv"),("all files", "*.*")))
    print(top.savefile)

    conn = pypyodbc.connect(
        r"Driver={Microsoft Access Driver (*.mdb)};" +
        r"Dbq=" + diag.filename
    )

    cur = conn.cursor()

    # SQL command to get the sequences from the Ovarian Cancer table
    command = cur.execute(querycode)

    with open(top.savefile + ".csv", "a") as myfile:
        myfile.write(",".join(header) + '\n')
        for row in command:
            row = list(map(str,row))
            myfile.write(",".join(row) + '\n')

    finish.config(text = "Query Finished", fg = 'blue')

    cur.close()
    conn.close()
    sys.stdout.close()

whereallele = Label(top, text = "Allele:")
entryallele = Entry(top, bd = 5)
whereallele.pack()
entryallele.pack()

whereantigen = Label(top, text = "Antigen Name:")
entryantigen = Entry(top, bd = 5)
whereantigen.pack()
entryantigen.pack()

sequencecheck = Checkbutton(top, text = "Query Sequences", variable = seqon, onvalue = 1, offvalue = 0)
allelecheck = Checkbutton(top, text = "Query Alleles", variable = alleleon, onvalue = 1, offvalue = 0)
antigencheck = Checkbutton(top, text = "Query Antigens", variable = antigenon, onvalue = 1, offvalue = 0)
affinitycheck = Checkbutton(top, text = "Query Affinity", variable = affinityon, onvalue = 1, offvalue = 0)
lengthcheck = Checkbutton(top, text = "Query Length", variable = lengthon, onvalue = 1, offvalue = 0)
poscheck = Checkbutton(top, text = "Query Position", variable = poson, onvalue = 1, offvalue = 0)

sequencecheck.pack()
allelecheck.pack()
antigencheck.pack()
affinitycheck.pack()
lengthcheck.pack()
poscheck.pack()

querybutton = Button(top, text = "Run Query", command = access)
querybutton.pack(pady= 10)

finish = Label(top, text = "")
finish.pack()

top.mainloop()
