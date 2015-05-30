from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, cm, letter, inch
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, Table, TableStyle
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_CENTER
from reportlab.lib import colors
from reportlab.lib.units import mm
from io import BytesIO

buffer = BytesIO()
file = open('Employee_Survey.pdf', 'w+')

width, height = letter
styles = getSampleStyleSheet()
styleN = styles["BodyText"]
styleN.alignment = TA_LEFT

styleBH = styles["Normal"]
styleBH.alignment = TA_CENTER
styleBH.fontSize = 7

def coord(x, y, unit=1):
    x, y = x * unit, height -  y * unit
    return x, y

def drawGrid(canvas):
    GRID_COLOR = (0.5, 0.5, 1)
    INTERVAL = 5 * mm

    numv = int((width - ((4 + 4) * mm)) / INTERVAL)
    numh = int((height - (16 + 3) * mm) / INTERVAL)
    voff = (width - numv * INTERVAL) / 2.0
    hoff = 16 * mm

    # Draw
    c = canvas
    c.setLineWidth(0.15)
    c.setStrokeColorRGB(*GRID_COLOR)
    for xi in xrange(numv + 1):
        c.line(xi * INTERVAL + voff, hoff,
               xi * INTERVAL + voff, hoff + numh * INTERVAL)
    for yi in xrange(numh + 1):
        c.line(voff, hoff + yi * INTERVAL,
               voff + numv * INTERVAL, hoff + yi * INTERVAL)
    #c.circle(width / 2, 5 * mm, 0.7, fill=1)
    c.save()

# Headers
perfMgmtHeader = Paragraph('''<b>Perfomance Managment</b>''', styleBH)
totalNheader = Paragraph('''<b>Total N''', styleBH)
percentResp = Paragraph('''<b>Percent Responding''', styleBH)
percentFavHeader = Paragraph('''<b>% Fav''', styleBH)
percentDistHeader = Paragraph('''<b>% Distribution''', styleBH)
meanHeader = Paragraph('''<b>Mean''', styleBH)

data= [[perfMgmtHeader, totalNheader,percentResp, percentFavHeader, percentDistHeader,meanHeader],
       ['00', '01', '02', '03', '04'],
       ['10', '11', '12', '13', '14'],
       ['20', '21', '22', '23', '24'],
       ['30', '31', '32', '33', '34']]

#Style the table
t = Table(data,style=[
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ('FONTSIZE',(0, 0), (-1, -1),7)
])

# Fixed column widths
t._argW[0] = 2*inch #Perf Mgmt
t._argW[1] = .43*inch #Total N
t._argW[2] = 2.5*inch # % Resp 1
t._argW[3] = .5*inch # % Fav
t._argW[4] = 1.5*inch # % Dist
t._argW[5] = .43*inch # Mean

c = canvas.Canvas(buffer, pagesize=A4)
t.wrapOn(c, width, height)
t.drawOn(c, *coord(5, 30, mm))
drawGrid(c)

#Write the buffer to disk and close the file
pdf = buffer.getvalue()
file.write(pdf)
buffer.close()

file.close()
