def draw_grid(canvas):
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
             {'name': 'sample 1', 'total_n': 12, 'responding':[20, 30, 50],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 2', 'total_n': 16, 'responding':[10,50,40],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 3', 'total_n': 12, 'responding':[20, 30, 50],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 4', 'total_n': 16, 'responding':[10,50,40],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 5', 'total_n': 12, 'responding':[20, 30, 50],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 6', 'total_n': 16, 'responding':[10,50,40],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 7', 'total_n': 12, 'responding':[20, 30, 50],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
             {'name': 'sample 8', 'total_n': 16, 'responding':[10,50,40],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
            ],
         },
         {
         'Tenure': [
             {'name': 'sample 1', 'total_n': 12, 'responding':[15, 15, 70],'distrobution': [2, 5, 26, 33, 34],'mean': 3.93,'favorable': 68},
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