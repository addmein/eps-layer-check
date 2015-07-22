from win32com.client import Dispatch
from win32com.client import constants
import win32com, os

ai = win32com.client.gencache.EnsureDispatch("Illustrator.Application.CS5.1")

def open(filename):
    ai.Open(filename)
    
def close():
    ai.Application.ActiveDocument.Close(constants.aiDoNotSaveChanges)
    
def list_layers():
    list = []
    for l in ai.ActiveDocument.Layers:
        list.append(l.Name)
    return list
        
folder = "C:\\temp\\general"

def check_layers(filename):
    open(filename)
    a = list_layers()
    if "TEXT" in a:
        return True
    else:
        return False
    close()

def check_files(folder):
    print "Checking files for TEXT layer..."
    print "=" * 32
    list = []
    for root, dirs, files in os.walk(folder):
        for file in files:
            fp = os.path.join(folder, file)
            if not check_layers(fp):
                list.append(fp)
            print "%s checked." %file
    if len(list) > 0:
        print "=" * 150
        print "There are %s files without the TEXT layer." %len(list)
        print "=" * 150
        print "\nList of files that need to be modified:"
        print "-" * 39
        for file in list:
            print file
    else:
        print "All the files contain the TEXT layer."
        
    
def exit_app():
    ai.Application.Quit()

check_files(folder)
exit_app()

print "=" * 100
print "TASK COMPLETED"