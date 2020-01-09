import pandas as pd

from datetime import timedelta
import datetime
import os
import math
import xlsxwriter

def trade_graph():

    year = '2019'
    graph = open(year+'_Graph.html', 'w')
    files = os.listdir('files')

    month = ['JAN','FEB', 'MAR', 'APR', 'MAY', 'JUNE', 'JULY', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']

    print('<!DOCTYPE html>', file = graph)
    print("<html lang ='en' dir='ltr'>", file = graph)
    print('<head><meta charset="utf-8">', file = graph)
    print('<script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>', file = graph)
    print('<link href="https://fonts.googleapis.com/css?family=Nanum+Gothic&display=swap" rel="stylesheet">', file = graph)
    print('<script type="text/javascript">', file = graph)

    print("google.charts.load('current', {'packages':['corechart']});", file = graph)
    print('google.charts.setOnLoadCallback(drawMultSeries);', file = graph)

    print('function drawMultSeries() {', file = graph)

    print('var data = google.visualization.arrayToDataTable([', file = graph)
    print("['Month', 'Within 5 days', {role: 'annotation'}, 'After 5 days', {role: 'annotation'}],", file = graph)
    print('\n')


    workbook = xlsxwriter.Workbook(year+"_Report.xlsx")
    worksheet = workbook.add_worksheet()

    worksheet.set_column('A:A', 15)
    worksheet.set_column("B:D", 30)
    bolds = workbook.add_format({'bold': True, 'font_size':18, 'border': 1})
    bold = workbook.add_format({'bold': True, 'align':'center', 'bg_color':'#A9A9A9', 'border': 1})
    border = workbook.add_format({'border': 1})

    worksheet.merge_range('A1:D1', 'UNICEF Country Offices FX Trade Delivery Days 01 JAN - 31 DEC '+year, bolds)

    worksheet.write('A2', 'Month', bold)
    worksheet.write('B2', 'Delivery within 5 days', bold)
    worksheet.write('C2', 'Delivery after 5 days', bold)
    worksheet.write('D2', 'Total number of transactions', bold)



    sum = 0
    row_record = 3

    for m in month:

        for f in files:

            if f[:3] == '.DS':

                print('DS File Store')

            else:

                TradeData = pd.read_excel("files/"+f, sheet_name = 0, header = None, skiprows=3)

                greentransactions = 0
                redtransactions = 0

                for days in TradeData[10]:

                    if isinstance(days, str) == True:
                        if int(days[:-5]) < 6:

                            greentransactions += 1

                        else:
                            redtransactions += 1
                if m == f[:-16]:

                    print(f[:-16], 'Within 5 days- transactions equals: ', greentransactions)
                    print(f[:-16], 'Greater than 5 days- transactions equals: ', redtransactions)
                    print('\n')

                    print('["',f[:-16],'",', greentransactions,',"', "{:.0f}".format( greentransactions/(greentransactions+redtransactions)*100 ),'%",', redtransactions,',"',  "{:.0f}".format( redtransactions/(greentransactions+redtransactions)*100 ),'%"],', file = graph)
                    worksheet.write('A'+str(row_record), m, border)
                    worksheet.write('B'+str(row_record), greentransactions, border)
                    worksheet.write('C'+str(row_record), redtransactions, border)
                    worksheet.write('D'+str(row_record), (redtransactions+greentransactions), border)


                    sum += (redtransactions+greentransactions)
        row_record+=1

    worksheet.write('C'+str(row_record), "TOTAL", bold)
    worksheet.write('D'+str(row_record), "{:,.2f}".format(sum), bold)
    worksheet.merge_range('B'+str(row_record+1)+':D'+str(row_record+1), "Compiled by: Louisa Tinga - Treasury Unit")
    print(']);', file = graph)

    print('var options = {', file = graph)
    print("title: 'UNICEF Country Offices FX Trade Delivery Days 01 JAN - 31 DEC 2019',", file = graph)
    print("chartArea: {width: '50%'},", file = graph)
    print('hAxis: {', file = graph)
    print("title: 'Number of Transactions',", file = graph)
    print("minValue: 0", file = graph)
    print(' },', file = graph)
    print('vAxis: {', file = graph)
    print("title: 'Months'", file = graph)
    print("}", file = graph)
    print('};', file = graph)
    print("var chart = new google.visualization.BarChart(document.getElementById('chart_div'));", file = graph)
    print("chart.draw(data, options);", file = graph)
    print("}", file = graph)
    print('</script>', file = graph)
    print('<style type="text/css">', file = graph)
    print("h2{", file = graph)
    print("font-family: 'Nanum Gothic', sans-serif;", file = graph)
    print("color:black;", file = graph)
    print("}", file = graph)
    print("h3{", file = graph)
    print("font-family: 'Nanum Gothic', sans-serif;", file = graph)
    print("color:grey;", file = graph)
    print('}', file = graph)
    print("td{ font-family: 'Nanum Gothic', sans-serif; }",file = graph)
    print("strong{ font-family: 'Nanum Gothic', sans-serif; }",file = graph)

    print("</style>", file = graph)
    print("</head>", file = graph)
    print('<body>', file = graph)
    #print('<h2>FX Trades 1-',calendar.monthrange(int(year), month_dict[getmonth])[1], getmonth, year,' Days Delivery To CO</h2>', file = graph)

    print('<div id="chart_div" style="width: 1600px; height: 1200px;"></div>', file = graph)

    print('<table style="border:none">', file = graph)
    print('<tr><td><strong>Key</strong></td><td></td></tr>', file = graph)

    print('<tr><td><div style="display:inline-block; width:50px; height:50px; background-color:blue"></div></td><td>FX Trades delivered to CO within 5 days</td>',file = graph)

    print('<tr><td><div style="display:inline-block; width:50px; height:50px; background-color:red"></div></td><td>FX Trades delivered to CO after 5 days onwards</td>',file = graph)
    print('</table>', file = graph)
    print('<p><strong>Compiled by: Louisa Tinga - Treasury Unit</strong></p>', file = graph)
    print('</body></html>', file = graph)

    workbook.close()
    graph.close()

        #for delete in files:


        #    os.remove('files/'+delete)
