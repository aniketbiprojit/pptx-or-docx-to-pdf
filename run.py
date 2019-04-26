from os import walk
import comtypes.client
import os

def PPTtoPDF(inputFileName, outputFileName, formatType=32):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    if outputFileName[-3:] != 'pdf':
        outputFileName = outputFileName + ".pdf"
    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType)  # formatType = 32 for ppt to pdf
    deck.Close()
    print('done')
    powerpoint.Quit()

search_dir = "./"
os.chdir(search_dir)
files = filter(os.path.isfile, os.listdir(search_dir))
files = [os.path.join(search_dir, f) for f in files] # add path to each file
files.sort(key=lambda x: os.path.getmtime(x))

# for (dirpath, dirnames, filenames) in walk('./'):
#     files.extend(filenames)
#     break
files.reverse()
print(files)
for file in files:
    if(file[-4:] == 'pptx'):
        print(file)
        print()
        PPTtoPDF(os.getcwd()+'\\'+file, os.getcwd()+'.\\'+file[:-4]+'pdf')
