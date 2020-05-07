import os
from zipfile import ZipFile
import glob
import shutil
import openpyxl as xl



levels={

    'INTRO':'INTROTEMPLATE.pptx',
    'BOOK1':'BOOK1TEMPLATE.pptx'

}




user_home=os.path.expanduser('~')

file = xl.load_workbook(os.path.abspath(glob.glob('Template.xlsx')[0]),data_only=True)
print(file)
print('Workbook loaded')
sheets = file.worksheets
sheet = sheets[0]
rows = sheet.rows
rows = list(rows)
row=[row.value for row in rows[0]]
# print(row)


def MassCreatePresentation():
    for i in range(0,len(rows)-1):
        level=rows[i+1][1].value[0:rows[i+1][1].value.index(' ')]
        template = os.path.abspath(glob.glob(levels[level])[0])
        print('Template Found')


        print(template)
        print('Copying file as ZIP')
        shutil.copy(template,'Operating/inwork.zip')

        print('Extracting Presentation')
        zip = ZipFile('Operating/inwork.zip')
        zip.extractall('Operating/')
        print('Extraction Successful')

        print('Getting and sorting Slide files')
        slideFiles = glob.glob('Operating/ppt/slides/*.xml')
        slideFiles.sort()
    # for element in slideFiles:


        for element in slideFiles:
            # os.system('clear')
            print('Operating on: '+os.path.basename(element))
            with open(element,'r+') as file:
                lines =file.read()

                while '&lt;&lt;' in lines:
                    tag =lines[lines.index('&lt;&lt;'):lines.index('&gt;&gt;')+8]
                    strippedtag=lines[lines.index('&lt;&lt;')+8:lines.index('&gt;&gt;')]
                    print('indicator found: '+strippedtag)

                    lines= lines.replace(tag,str(rows[i+1][row.index(strippedtag)].value))

                    print('Replaced')
                file.seek(0)
                file.truncate(0)
                file.write(lines)
        finalPath='MassProduced/'+str(rows[i+1][0].value)+' '+str(rows[i+1][1].value)+' '+str(rows[i+1][2].value)
        with ZipFile('MassProduced/'+str(rows[i+1][0].value)+' '+str(rows[i+1][1].value)+' '+str(rows[i+1][2].value)+'.zip','w') as zip:
            os.remove(glob.glob('Operating/*.zip')[0])
            # shutil.make_archive('MassProduced/'+str(rows[i+1][0].value)+' '+str(rows[i+1][1].value)+' '+str(rows[i+1][2].value),'zip','Operating/','')


            for folder, subfolder,files in os.walk('Operating/'):
                for file in files:
                    print(file)
                    print(folder)
                    filePath = os.path.join(folder,file)
                    print('FILEPATH IS: ' + filePath)
                    print('USEFULPATH IS: '+filePath[filePath.index('/')+1:])
                    zip.write(filePath,filePath[filePath.index('/')+1:])

        os.rename(finalPath+'.zip',finalPath+'.pptx')

            # break
            # for line in lines:
            #     new_line=line
            #     while '&lt;&lt;' in new_line:
            #         tag =new_line[new_line.index('&lt;&lt;'):new_line.index('&gt;&gt;')+8]
            #         print('indicator found: '+tag)
            #
            #         new_line= new_line.replace(tag,'NEWINDICATOR')
            #
            #         print('Replaced')



    return





MassCreatePresentation()
