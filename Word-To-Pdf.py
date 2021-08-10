import os
import win32com.client

print('''

##      ##  #######  ########  ########          ########  ########  ######## 
##  ##  ## ##     ## ##     ## ##     ##         ##     ## ##     ## ##       
##  ##  ## ##     ## ##     ## ##     ##         ##     ## ##     ## ##       
##  ##  ## ##     ## ########  ##     ## ####### ########  ##     ## ######   
##  ##  ## ##     ## ##   ##   ##     ##         ##        ##     ## ##       
##  ##  ## ##     ## ##    ##  ##     ##         ##        ##     ## ##       
 ###  ###   #######  ##     ## ########          ##        ########  ##       

''')

print('Made By:-  Hacker--Rohan Raj')

input('Press Enter To Continue')

wdFormatPDF = 17

pathtoconvert = os.getcwd()

for root, dirs, files in (os.walk(pathtoconvert)):
    for f in files:


        if  f.endswith(".doc")  or f.endswith(".odt") or f.endswith(".rtf") or f.endswith(".docx"):
            try:
                print(f)
                in_file=os.path.join(root,f)
                word = win32com.client.Dispatch('Word.Application')
                word.Visible = False
                doc = word.Documents.Open(in_file)
                doc.SaveAs(os.path.join(root,f[:-4]), FileFormat=wdFormatPDF)
                doc.Close()
                word.Quit()
                word.Visible = True
                print ('done')
                os.remove(os.path.join(root,f))
                pass
            except:
                print('could not open')
                # os.remove(os.path.join(root,f))
        elif f.endswith(".docx") or f.endswith(".dotm") or f.endswith(".docm"):
            try:
                print(f)
                in_file=os.path.join(root,f)
                word = win32com.client.Dispatch('Word.Application')
                word.Visible = False
                doc = word.Documents.Open(in_file)
                doc.SaveAs(os.path.join(root,f[:-5]), FileFormat=wdFormatPDF)
                doc.Close()
                word.Quit()
                word.Visible = True
                print ('done')
                os.remove(os.path.join(root,f))
                pass
            except:
                print('could not open')
                # os.remove(os.path.join(root,f))
        else:
            del dirs[:]
            pass