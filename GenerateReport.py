#Developer: Dean Hutton
#Project: ReportLabs Report
#Date: 5/31/15

#Define imports
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, Table
from reportlab.lib.enums import TA_LEFT, TA_RIGHT,TA_CENTER
from reportlab.platypus import SimpleDocTemplate
from reportlab.lib.units import mm
from reportlab.graphics.shapes import *
from io import BytesIO
smallInputExample = {
     'summary': {
     'total_n': 114,
     'responding': [68, 26, 6],
     'favorable': 68,
     'distrobution': [2, 5, 26, 33, 34],
     'mean': 3.93,
     },
     'demographics': [
         {
         'Locations': [
             {'name': 'dean 1', 'total_n': 12, 'responding':[20,30,50],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 2', 'total_n': 16, 'responding':[33,33,33],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
            ],
         },
         {
         'People Managment': [
             {'name': 'sample 1', 'total_n': 12, 'responding':[20, 30, 50],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 2', 'total_n': 16, 'responding':[10,50,40],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},

            ],
         },
         {
         'Tenure': [
             {'name': 'sample 1', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 2', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
            ],
         },
     ],
 }
#Use this list later on to SPAN on location, People MGMT, Tenure Labels
labelRowList = []

#BytesIO buffer as well as output file to save buffer too. Eventually output file will be replace with HTTP request
buffer = BytesIO()
file = open('Employee_Survey.pdf', 'w+')


#Define styles and page properties
width, height = letter
styles = getSampleStyleSheet()
styleN = styles['BodyText']
styleN.alignment = TA_LEFT

headerStyle = styles['Normal']
headerStyle.alignment = TA_CENTER
headerStyle.fontSize = 7
headerStyle.fontName = 'Helvetica-Bold'

rightAlignedStyle = styles['Normal']
rightAlignedStyle.alignment = TA_RIGHT
rightAlignedStyle.fontSize = 7
rightAlignedStyle.fontName = 'Helvetica'

leftAlignedStyle = styles['Normal']
leftAlignedStyle.alignment = TA_LEFT
leftAlignedStyle.fontSize = 7
leftAlignedStyle.fontName = 'Helvetica'

'''Function that is used to generate charts. The methodolgy is to create 3 rectangles and combine them. I used
the length of the cell which is 60mm and calculated percentages off that
'''
def build_percent_responding_rectangles(percentFavorable,percentNeutral,percentUnfavorable,header=False):

    #Convert the ints to floats so we can caclulate percentages
    percentFavorableFloat = float(percentFavorable)/100
    percentNeutralFloat = float(percentNeutral)/100
    percentUnfavorableFloat = float(percentUnfavorable)/100
    #If its the header than give it an offset so we can fit the Percent Responding name
    if header:
        d = Drawing(0*mm, 10*mm)

    else:
        d = Drawing(0*mm, 0*mm)

    #Used totalRectLen to calculate percentages
    totalRectangleLength = 60
    rectangleHeight = 4
    rectangleYposition = 0
    rectangleTextYposition = 1
    rectangleTextHeight = 2.5

    #Offset that is used to account for two letters in percent e.g. 11 or 74
    rectangleTextXPositionOffset = 1.25

    #Favorable Rect Creation. Align the text to 50% of the total rect length
    favorableXposition = 0
    favorableRectWidth = percentFavorableFloat*totalRectangleLength
    if header:
        favorableRectTextXPosition = favorableXposition + 1
        favorableRectText = '% Favorable'
    else:
        favorableRectTextXPosition = (favorableXposition +favorableRectWidth)/2 - rectangleTextXPositionOffset
        favorableRectText = str(percentFavorable)

    d.add(Rect(favorableXposition*mm, rectangleYposition*mm, favorableRectWidth*mm, rectangleHeight*mm, fillColor=colors.green,strokeColor=colors.green))
    d.add(String(favorableRectTextXPosition*mm, rectangleTextYposition*mm, favorableRectText, fontSize=rectangleTextHeight*mm,fontName='Helvetica'))

    #Neutral Rect creation
    neutralXposition = favorableRectWidth
    neutralRectWidth = percentNeutralFloat*totalRectangleLength
    if header:
        neutralRectTextXPosition = neutralXposition + 1
        neutralRectText = '% Neutral'
    else:
        neutralRectTextXPosition = (neutralXposition + neutralRectWidth + favorableXposition+favorableRectWidth)/2 - rectangleTextXPositionOffset
        neutralRectText = str(percentNeutral)

    d.add(Rect(neutralXposition*mm, rectangleYposition*mm, neutralRectWidth*mm, rectangleHeight*mm, fillColor=colors.yellow, strokeColor=colors.yellow))
    d.add(String(neutralRectTextXPosition*mm, rectangleTextYposition*mm, neutralRectText, fontSize=rectangleTextHeight*mm,fontName='Helvetica'))

    #Unfavorable Rect
    unfavorableXposition = neutralRectWidth + favorableRectWidth
    unfavorableRectWidth = (percentUnfavorableFloat*totalRectangleLength)
    if header:
        unfavorableRectTextXPosition = unfavorableXposition + 1
        unfavorableRectText = '% Unfavorable'
    else:
        unfavorableRectTextXPosition = (unfavorableXposition + unfavorableRectWidth + neutralXposition + neutralRectWidth)/2 - rectangleTextXPositionOffset
        unfavorableRectText = str(percentUnfavorable)

    d.add(Rect(unfavorableXposition*mm, rectangleYposition*mm, unfavorableRectWidth*mm, rectangleHeight*mm, fillColor=colors.red, strokeColor=colors.red))
    d.add(String(unfavorableRectTextXPosition*mm,rectangleTextYposition*mm, unfavorableRectText, fontSize=rectangleTextHeight*mm,fontName='Helvetica'))

    if header:
        #Percent Responding Heading text
        percentResponseTextXposition = 15
        percentResponseTextYposition = 5
        percentResponseText = 'Percent Responding'
        d.add(String(percentResponseTextXposition*mm,percentResponseTextYposition*mm, percentResponseText, fontSize=7, fontName='Helvetica-Bold'))
    return d

#Function that adds spacing in the % Distrobution column
def percent_distrobution_space_equalizer(distrobutionList):
    distroString = ''
    fourHTMLspaces = ' &nbsp &nbsp '
    for item in distrobutionList:
        distroString = distroString + str(item) + fourHTMLspaces
    return distroString

#Function that creates paragraph objects for texts so that they can be styled
def create_paragraph_and_chart_row(parameterDict, label = '',bold = False):

    if bold:
        labelParagraph = Paragraph('''<b>%s'''%label, rightAlignedStyle)
    else:
        labelParagraph = label

    TotalNParagraph = Paragraph('''%s'''%parameterDict.get('total_n'), rightAlignedStyle)
    Chart = build_percent_responding_rectangles(parameterDict.get('responding')[0],parameterDict.get('responding')[1],parameterDict.get('responding')[2],header=False)
    FavorableParagraph =  Paragraph('''%s'''%parameterDict.get('favorable'), rightAlignedStyle)
    Distrobution = percent_distrobution_space_equalizer(parameterDict.get('distrobution'))
    DistParagraph = Paragraph('''%s'''%Distrobution, rightAlignedStyle)
    MeanParagraph =  Paragraph('''%s'''%parameterDict.get('mean'), rightAlignedStyle)

    return [labelParagraph, TotalNParagraph, Chart, FavorableParagraph, DistParagraph, MeanParagraph]

def populate_chart_data(input):
    data = []
    # Create a lable row counter to keep track of which rwos the lables are in so we can SPAN them later on. Start at 1 to account for table header
    labelRowCounter = 1
    summaryData = input.get('summary')
    if summaryData:
        summaryList = create_paragraph_and_chart_row(summaryData,'Overall Company',True)
        data.append(summaryList)

    demographicData = input.get('demographics')

    #Iterate through the different demographics and create rows for them
    for demoDicts in demographicData:
        headerKey = demoDicts.keys()[0]
        headerParagraph = Paragraph('''<b>%s'''%headerKey, rightAlignedStyle)
        labelRowCounter += 1
        labelRowList.append(labelRowCounter)
        data.append([headerParagraph,'','','','',''])
        for demoDict in demoDicts[headerKey]:
            sampleName = demoDict.get('name','')
            demoList = create_paragraph_and_chart_row(demoDict,sampleName,False)
            data.append(demoList)
            labelRowCounter += 1
    return data

#Build the header for the report
percentRespondingHeader = build_percent_responding_rectangles(33, 33, 33, header=True)
chart = build_percent_responding_rectangles(20, 30, 50, header=False)
perfMgmtHeader = Paragraph('''<b>Perfomance Managment</b>''', headerStyle)
totalNheader = Paragraph('''<b>Total N''', headerStyle)
percentResp = Paragraph('''<b>Percent Responding''', headerStyle)
percentFavHeader = Paragraph('''<b>% Fav''', headerStyle)
percentDistHeader = Paragraph('''<b>% Distribution  &nbsp &nbsp &nbsp &nbsp  &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp 1 &nbsp &nbsp  2 &nbsp &nbsp 3 &nbsp &nbsp 4 &nbsp &nbsp 5''', headerStyle)
meanHeader = Paragraph('''<b>Mean''', headerStyle)
reportHeader = [perfMgmtHeader, totalNheader,percentRespondingHeader, percentFavHeader, percentDistHeader,meanHeader]

#Populate the chart with data. This is where the inputDict will be passed in
data = populate_chart_data(smallInputExample)

reportData = []
reportData.append(reportHeader)

#Populate the reportData with rows
for d in data:
    reportData.append(d)

#Style the table
tableStyle = [
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ('FONTSIZE',(0, 0), (-1, -1), 7),
    ('ALIGN', (0, 3), (0, -1), 'RIGHT'),
]

#Add additional styling to the table so that lables such as locations, People Management, Tenure
for labelRow in labelRowList:
    tup = ('SPAN',(0,labelRow),(-1,labelRow))
    tableStyle.append(tup)

#Set the table with data and style
t = Table(reportData,style=tableStyle)
t.hAlign = 'LEFT'

#Fixed column widths
t._argW[0] = 50.8*mm # Perf Mgmt
t._argW[1] = 12.7*mm # Total N
t._argW[2] = 63.5*mm # % Resp
t._argW[3] = 12.7*mm # % Fav
t._argW[4] = 38.1*mm # % Dist
t._argW[5] = 10.922*mm # Mean

#Build the document. Use SimpleDocTemplate to make page breaks and table alignment equal for different data sets
doc = SimpleDocTemplate(buffer,pagesize=letter,rightMargin=30,leftMargin=30,topMargin=30,bottomMargin=30)
doc.build([t])

#Write the buffer to disk and close the file
pdf = buffer.getvalue()
file.write(pdf)
buffer.close()
file.close()
