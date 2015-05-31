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
             {'name': 'sample 1', 'total_n': 12, 'responding':[20,30,50],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
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

largeInputExample = {
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
             {'name': 'sample 1', 'total_n': 12, 'responding':[20,30,50],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 2', 'total_n': 16, 'responding':[33,33,33],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 3', 'total_n': 12, 'responding':[20, 30, 50],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 4', 'total_n': 16, 'responding':[10,50,40],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 5', 'total_n': 12, 'responding':[20, 30, 50],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 6', 'total_n': 16, 'responding':[10,50,40],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
            ],
         },
         {
         'People Managment': [
             {'name': 'sample 1', 'total_n': 12, 'responding':[70, 15, 15],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 2', 'total_n': 16, 'responding':[10,50,40],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 3', 'total_n': 12, 'responding':[20, 30, 50],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 4', 'total_n': 16, 'responding':[10,50,40],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 5', 'total_n': 12, 'responding':[50, 25, 25],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 6', 'total_n': 16, 'responding':[25,50,25],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 7', 'total_n': 12, 'responding':[20, 30, 50],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 8', 'total_n': 16, 'responding':[10,50,40],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
            ],
         },
         {
         'Tenure': [
             {'name': 'sample 1', 'total_n': 12, 'responding':[70, 10, 20],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 2', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 3', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 4', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 5', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 6', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 7', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 8', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 9', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 10', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 11', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 12', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 13', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 14', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 15', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 16', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 17', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 18', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 19', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 20', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 21', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 22', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 23', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 24', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 25', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 26', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 27', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 28', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 29', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 30', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 31', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 32', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 33', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 35', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 36', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 37', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 38', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 39', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 40', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 41', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 42', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 43', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 44', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 45', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 46', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 47', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 48', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 49', 'total_n': 16, 'responding':[10,30,60],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
            ],
         },
     ],
 }

#Function that is used to generate charts. The methodolgy is to create 3 rectangles and put them tohe
def build_percent_responding_rectangles(percentFavorable,percentNeutral,percentUnfavorable,header=False):

    #Header Rectangles
    percentFavorableFloat = float(percentFavorable)/100
    percentNeutralFloat = float(percentNeutral)/100
    percentUnfavorableFloat = float(percentUnfavorable)/100
    if header:
        d = Drawing(0*mm, 10*mm)

    else:
        d = Drawing(0*mm, 0*mm)

    totalRectangleLength = 60
    rectangleHeight = 4
    rectangleYposition = 0
    rectangleTextYposition = 1
    rectangleTextHeight = 2.5

    #Offset that is used to account for two letters in percent e.g. 11 or 74
    rectangleTextXPositionOffset = 1.25

    #Favorable Rect Creation
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

    #Neutral Rect
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

def percent_distrobution_space_equalizer(distrobutionList):
    distroString = ''
    fourHTMLspaces = " &nbsp &nbsp "
    for item in distrobutionList:
        distroString = distroString + str(item) + fourHTMLspaces

    return distroString

def create_paragraph_and_chart_row(parameterDict, label = '',bold = False):

    if bold:
        labelParagraph = Paragraph('''<b>%s'''%label, rightAlignedStyle)
    else:
        labelParagraph = label

    TotalNParagraph = Paragraph('''%s'''%parameterDict.get("total_n"), rightAlignedStyle)
    Chart = build_percent_responding_rectangles(parameterDict.get('responding')[0],parameterDict.get('responding')[1],parameterDict.get('responding')[2],header=False)
    FavorableParagraph =  Paragraph('''%s'''%parameterDict.get("favorable"), rightAlignedStyle)
    Distrobution = percent_distrobution_space_equalizer(parameterDict.get('distrobution'))
    DistParagraph = Paragraph('''%s'''%Distrobution, rightAlignedStyle)
    MeanParagraph =  Paragraph('''%s'''%parameterDict.get("mean"), rightAlignedStyle)

    return [labelParagraph, TotalNParagraph, Chart, FavorableParagraph, DistParagraph, MeanParagraph]

def populate_chart_data(input):
    data = []
    # Create a lable row counter to keep track of which raws are lables so we can SPAN them later on. Start at 1 to account for table header
    labelRowCounter = 1
    summaryData = input.get('summary')
    if summaryData:
        summaryList = create_paragraph_and_chart_row(summaryData,"Overall Company",True)
        data.append(summaryList)

    demographicData = input.get('demographics')

    for demoDicts in demographicData:
        headerKey = demoDicts.keys()[0]
        headerParagraph = Paragraph('''<b>%s'''%headerKey, rightAlignedStyle)
        labelRowCounter += 1
        labelRowList.append(labelRowCounter)
        data.append([headerParagraph,"","","","",""])
        for demoDict in demoDicts[headerKey]:
            sampleName = demoDict.get("name","")
            demoList = create_paragraph_and_chart_row(demoDict,sampleName,False)
            data.append(demoList)
            labelRowCounter += 1

    return data


percentRespondingHeader = build_percent_responding_rectangles(33, 33, 33, header=True)
chart = build_percent_responding_rectangles(20, 30, 50, header=False)
perfMgmtHeader = Paragraph('''<b>Perfomance Managment</b>''', headerStyle)
totalNheader = Paragraph('''<b>Total N''', headerStyle)
percentResp = Paragraph('''<b>Percent Responding''', headerStyle)
percentFavHeader = Paragraph('''<b>% Fav''', headerStyle)
percentDistHeader = Paragraph('''<b>% Distribution  &nbsp &nbsp &nbsp &nbsp  &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp 1 &nbsp &nbsp  2 &nbsp &nbsp 3 &nbsp &nbsp 4 &nbsp &nbsp 5''', headerStyle)
meanHeader = Paragraph('''<b>Mean''', headerStyle)
reportHeader = [perfMgmtHeader, totalNheader,percentRespondingHeader, percentFavHeader, percentDistHeader,meanHeader]

data = populate_chart_data(largeInputExample)
chartData = []
chartData.append(reportHeader)


for d in data:
    chartData.append(d)

tableStyle = [
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ('FONTSIZE',(0, 0), (-1, -1), 7),
    ('ALIGN', (0, 3), (0, -1), 'RIGHT'),
]

for labelRow in labelRowList:
    tup = ('SPAN',(0,labelRow),(-1,labelRow))
    tableStyle.append(tup)

#Style the table


t = Table(chartData,style=tableStyle)
t.hAlign = "LEFT"


# Fixed column widths
t._argW[0] = 50.8*mm #Perf Mgmt
t._argW[1] = 12.7*mm #Total N
t._argW[2] = 63.5*mm # % Resp
t._argW[3] = 12.7*mm # % Fav
t._argW[4] = 38.1*mm # % Dist
t._argW[5] = 10.922*mm # Mean


doc = SimpleDocTemplate(buffer,pagesize=letter,rightMargin=30,leftMargin=30,topMargin=30,bottomMargin=30)
doc.build([t])

#Write the buffer to disk and close the file
pdf = buffer.getvalue()
file.write(pdf)
buffer.close()
file.close()
