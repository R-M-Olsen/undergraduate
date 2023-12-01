import csv, exifread, os, pyproj, subprocess, time
from easygui import *

# default variables
NoDataVal = -9999
LocalDTG = False
UTCoffset = False
GPSdatum = False
GPSposErr = False
outputFile = "coord_export_%s" % (time.strftime("%Y%m%d-%H%M%S"))
header = "fileName, lon, lat, alt"
outputList = []
CSVfields = ['fileName', 'lon', 'lat', 'alt']
CSVrows = []
excluded = ['The following files have been excluded from export:\n\n\tFile:\t\t\tReason:']
excludeNum = 0
calcPointStats = False
UTMzone = ''
# intro message
def intro():
#    introMsg = ('This program will extract exif GPS data from images and export it to a csv file for easy import to' +
#                'GIS\n\nJust follow the steps')
    introTtl = "Please Confirm"
    introMsg = '''
Simple Exif to CSV
	Written by R. Olsen
	November 2023

Table of Contents
	A: Usage
	B: What this program does
	C: Python packages used
	
A: Usage:
	1: Set output file name and noData value.
	2: Select any extra fields you would like top add to the output.
	3: Select the folder where your images are located.
	4: Review extracted data
	5: Select folder where you want the csv to be saved or cancel and restart

B: What this program does:
	• Extracts Exif data from .jpg, .png, .wav, .webp, and .tif files into a .csv file.
	• Exif tags extracted:
		- GPSLongitude
		- GPSLongitudeRef
		- GPSLatitude
		- GPSLatitudeRef
		- GPSAltitude
		- GPSAltitudeRef
		- DateTime (optional)
		- DateTimeDigitized (optional)
		- OffsetTime (optional)
		- GPSMapDatum (optional)
		- GPSHPositioningError (optional)

C: Python packages used:
	• csv
	• ExifRead (https://pypi.org/project/ExifRead/)
	• Easy GUI (https://pypi.org/project/easygui/)
	• os
	• subprocess
    '''
    if ccbox(introMsg, introTtl):     # show a Continue/Cancel dialog
        #pass  # user chose Continue
        getSettings()
    else:  # user chose Cancel
        exit('user cancelled at intro')

# get settings
def getSettings():
    global outputFile, NoDataVal
    setMsg = ('Please set your variables\n\nDefault values:\n\tFile Name... %s\n\tNoData......... %s\n\nSelecing \'Ok\' with' +
              'no entry will keep default.') % (outputFile, NoDataVal)
    setTtl = 'Settings'
    setNames = ["Output File Name", "NoData Value"]
    setValues = multenterbox(setMsg, setTtl, setNames)
    if setValues is None:
        exit('user cancelled at settings')
    while 1:
        if setValues != None:
            if setValues[0].strip() != "":
                outputFile = setValues[0].strip()
            if setValues[1].strip() != "":
                NoDataVal = setValues[1].strip()
            extra_columns()
            break # no problems found
        else:
            extra_columns()
            break

def extra_columns():
    global header,LocalDTG,UTCoffset,GPSdatum,GPSposErr
    # user select extra columns
    colSetMsg = ('Please select which extra fields you would like to include in the export.\n\nSelecting \'Cancel\' ' +
                 'has the same effect as selecting none\n\nNOTE: Images may not have the selected data, empty fields' +
                 'will show the selected NoData value (%s)') % (NoDataVal)
    colChoices = ["A:  Local time image was taken [LocalDTG]",
                  "B:  Timezone offset [UTCoffset]",
                  "C:  GPS datum [GPSdatum]",
                  "D:  GPS horizontal position error [GPSposErr]"
                  ]
    colSet = multchoicebox(colSetMsg, 'Field Selection', colChoices, None)
    while 1:
        if colSet is None:
            #bulk_extract()
            point_stats_ask()
        else:
            for i in range(len(colSet)):
                if colSet[i] == "A:  Local time image was taken [LocalDTG]":
                    LocalDTG = True
                    header += ', LocalDTG'
                    CSVfields.append('LocalDTG')
                if colSet[i] == "B:  Timezone offset [UTCoffset]":
                    UTCoffset = True
                    header += ', UTCoffset'
                    CSVfields.append('UTCoffset')
                if colSet[i] == "C:  GPS datum [GPSdatum]":
                    GPSdatum = True
                    header += ', GPSdatum'
                    CSVfields.append('GPSdatum')
                if colSet[i] == "D:  GPS horizontal position error [GPSposErr]":
                    GPSposErr = True
                    header += ', GPSposErr'
                    CSVfields.append('GPSposErr')
            #bulk_extract()
            point_stats_ask()

def folder_select():
    folder = diropenbox('Pick the folder where your images are stored.', 'Image folder selection',
                        default="./")
    return folder

def point_stats_ask():
    global calcPointStats
    ptStatBox = buttonbox('Do you want to calculate point statistics for similar points? This uses the Welzl\'s' +
                          'Algorithm to solve the smallest circle problem for points the same prefix.  For large' +
                          'numbers of points this may take longer than expected.\n\nPLEASE NOTE:\n\t• For point' +
                          'statistics to be calculated images of the same point must have three character prefixes' +
                          'that match.\n\t• At minimum of 3 points is needed.', 'Point Statistics',
        ['Calculate\nStatistics','Do not calculate\nstatistics'])
    if ptStatBox == 'Calculate\nStatistics':
        calcPointStats = True
    bulk_extract()

def get_UTM_zone():
    global UTMzone
    setMsg = ('Please enter the UTM zone where your images were take') % (outputFile, NoDataVal)
    setTtl = 'Settings'
    setNames = ["Output File Name", "NoData Value"]
    setValues = enterbox(setMsg, setTtl, setNames)
    if setValues is None:
        exit('user cancelled at settings')
    while 1:
        if setValues != None:
            if setValues[0].strip() != "":
                outputFile = setValues[0].strip()
            if setValues[1].strip() != "":
                NoDataVal = setValues[1].strip()
            extra_columns()
            break # no problems found
        else:
            extra_columns()
            break


def createLine(file, tags):
    global CSVrows, CSVfieldsm, excluded, excludeNum
    addColumns = []
    # Get exif longitude and convert to decimal degrees
    if 'GPS GPSLongitude'in tags:
        (LONdeg, LONmin, LONsec) = tags.get('GPS GPSLongitude').values
        GPSlon = LONdeg + (LONmin / 60) + (LONsec.decimal() / 3600)
        LONref = tags['GPS GPSLongitudeRef'].values
        if LONref == "W":
            GPSlon = (-1) * GPSlon
        GPSlonStr = int(GPSlon * 10 ** 7) / 10.0 ** 7  # limits decimal places to 7 without rounding
    else:
        GPSlonStr = str(NoDataVal)

    # Get exif latitude and convert to decimal degrees
    if 'GPS GPSLatitude' in tags:
        (LATdeg, LATmin, LATsec) = tags.get('GPS GPSLatitude').values
        GPSlat = LATdeg + (LATmin / 60) + (LATsec.decimal() / 3600)
        LATref = tags['GPS GPSLatitudeRef'].values
        if LATref == "S":
            GPSlat = (-1) * GPSlat
        GPSlatStr = int(GPSlat * 10 ** 7) / 10.0 ** 7  # limits decimal places to 7 without rounding
    else:
        GPSlatStr = str(NoDataVal)

    if GPSlatStr == str(NoDataVal): # or GPSlonStr == NoDataVal:
        exclude = '\n\t' + file + '\t\t\t' + 'No latitude or longitude '
        excluded.append(exclude)
        excludeNum += 1
        return None

    # get altitude
    ALTref = None
    ALTval = str(NoDataVal)
    if 'GPS GPSAltitudeRef' and 'GPS GPSAltitude' in tags:
        ALTref = tags['GPS GPSAltitudeRef'].values
        ALTval = tags['GPS GPSAltitude'].values[0]
    if ALTref == 1:  # 0 = Above sea level , 1 = Below sea level
        ALTval = -1 * ALTval

    fileRow = {'fileName':file, 'lon':str(GPSlonStr), 'lat': str(GPSlatStr), 'alt': str(ALTval)}

    # get date time group
    DTG = ''
    if LocalDTG == True:
        DTG = str(NoDataVal)
        if 'Image DateTime' in tags:
            #DTG = (', %s') % (tags['Image DateTime'].values)
            DTG = str(tags['Image DateTime'].values)
        else:
            if 'EXIF DateTimeDigitized' in tags:
                #DTG =  (', %s') % (tags['EXIF DateTimeDigitized'].values)
                DTG = str(tags['EXIF DateTimeDigitized'].values)
        addColumns.append(DTG)
        fileRow['LocalDTG'] = str(DTG)

    # get UTC offset
    UTC = ''
    if UTCoffset == True:
        UTC = str(NoDataVal)
        if 'EXIF OffsetTime' in tags:
            UTC = str(tags['EXIF OffsetTime'].values)
        addColumns.append(UTC)
        fileRow['UTCoffset'] = str(UTC)

    # get GPS datum
    datum = ''
    if GPSdatum == True:
        datum = str(NoDataVal)
        if 'GPS GPSMapDatum' in tags:
            datum = str(tags['GPS GPSMapDatum'].values)
        addColumns.append(datum)
        fileRow['GPSdatum'] = str(datum)

    # get UTC offset
    posErr = ''
    if GPSposErr == True:
        posErr = str(NoDataVal)
        if 'GPS GPSHPositioningError' in tags:
            posErr = str(tags['GPS GPSHPositioningError'].values)
        addColumns.append(posErr)
        fileRow['GPSposErr'] = str(posErr)

    # output values
    lineString = str(GPSlonStr) + ', ' + str(GPSlatStr) + ', ' + str(ALTval)

    # add extra columns
    for i in range(len(addColumns)):
        if addColumns[i] != '':
            lineString += ', ' + addColumns[i]

    CSVrows.append(fileRow)

    return file + ', ' + lineString

def bulk_extract():
    global outputList, header, excluded, excludeNum
    outputList.append(header)
    # open target folder
    folder = diropenbox('Pick the folder where your images are stored.','Image folder election', default="./")
    for imageName in os.listdir(folder):
        image = os.path.join(folder, imageName)
        shortName = imageName.split(".")[0]
        if os.path.isfile(image):
            if image.lower().endswith(('.tif', '.jpg', '.wav', '.webp')):
                OutStr = createLine(shortName, exifread.process_file(open(image, 'rb')))
            #outputList.append('\n' + shortName + ',' + OutStr)
                if OutStr != None:
                    outputList.append('\n' + OutStr)
            else:
                exclude = '\n\t' + shortName + '\t\t\t' + 'File is not .tif, .jpg, .wav, or .webp'
                excluded.append(exclude)
                excludeNum += 1
        else:
            exclude = '\n\t' + shortName + '\t\t\t' + 'Object is not a file'
            excluded.append(exclude)
            excludeNum += 1

    output_preview()

def writeCSV():
    global CSVrows,CSVfields,outputFile

    saveLoc = diropenbox('Pick the folder where you want the output saved.', 'Output folder selection',
                         default="./")
    filePath = str(saveLoc + '/' + outputFile + '.csv')
    #saveLoc = filesavebox('File Save As', 'File Save As', default=outputFile + ".csv")
    with open(filePath, 'w', newline='') as csvfile:
        # creating a csv writer object
        csvwriter = csv.DictWriter(csvfile, fieldnames=CSVfields)
        # writing the fields
        csvwriter.writeheader()
        # writing the data rows
        csvwriter.writerows(CSVrows)

    subprocess.run([os.path.join(os.getenv('WINDIR'), 'explorer.exe'),'/select,', os.path.normpath(filePath)])
    exit('normal end')

def output_preview():
    global NoDataVal,LocalDTG,UTCoffset,GPSdatum,GPSposErr,outputFile,header,CSVfields,CSVrows,outputList,excluded
    CSVpreviewMsg = ('CSV output preview\n\nPressing \'Ok\' will open a file save dialog.\n\nClicking \'Cancel\' ' +
                     'will open a window to quit or make changes and retry.')
    CSVpreviewTtl = 'CSV output preview'

    combinedOutput = []

    if excludeNum == 0:
        combinedOutput = outputList
    else:
        excluded.append('\n\n---------------------------- END EXCLUSIONS ----------------------------\n\n')
        combinedOutput = excluded + outputList

    reply = codebox(CSVpreviewMsg, CSVpreviewTtl, combinedOutput)
    if reply:
        writeCSV()
    else:
        lastBox = buttonbox('What do you want to do now?', 'End or restart.'
                            ,['Restart','Quit'],cancel_choice='Quit')
        if lastBox == 'Restart':
            NoDataVal = -9999
            LocalDTG = False
            UTCoffset = False
            GPSdatum = False
            GPSposErr = False
            outputFile = "coord_export_%s" % (time.strftime("%Y%m%d-%H%M%S"))
            header = "fileName, lon, lat, alt"
            outputList = []
            CSVfields = ['fileName', 'lon', 'lat', 'alt']
            CSVrows = []
            intro()
        else:
            exit('user cancelled at retry')
# Start script
intro()
