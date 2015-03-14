from pywinauto import application

# Code to handle the initial window that pops up regarding
# seperating commas, etc

excel = application.Application()

# Get the name of the file the fdv data was saved in to get the excel
# window to send events to
print "Input the file name without the .xls"
t = raw_input()
t = str(t)
complete_t = "Microsoft Excel - " + t
excel.connect_(title_re=complete_t)

# Set the tab to read "data"
excel[complete_t].ClickInput(coords=(104,1026), double=True)
excel[complete_t].TypeKeys("data")

# Open the macro menu
excel[complete_t].ClickInput(coords=(944,95))
# Click the macro title edit box
excel[complete_t].ClickInput(coords=(95,0), double=True)
excel[complete_t].TypeKeys("{DELETE}")

macros = ["V2791_FDV_D","V1900_FDV_A", "V1901_FDV_A"]
# Get the name of the macro to run
print "Type the index of the macro to run."
for x in range(len(macros)):
    print str(x) + ":  " + macros[x]
m = raw_input()
m = int(m)
m = macros[m]
          
excel[complete_t].TypeKeys(m)

# Click the run button
excel[complete_t].ClickInput(coords=(348,10))


