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


files = []
for (dirpath, dirnames, filenames) in walk('./'):
    files.extend(filenames)
    break

for file in files:
    if(file[-4:] == 'pptx'):
        print(file)
        print()
        PPTtoPDF(os.getcwd()+'\\'+file, os.getcwd()+'.\\'+file[:-4]+'pdf')
