import pandas
import xlrd
import plotly as py
py.tools.set_credentials_file(username='Yanhe05', api_key='JrXk8wZEc5P0X1LnlIe9')
import plotly.plotly as py

us_state = {
    'Alabama':'AL',
    'Alaska':'AK',
    'Arizona':'AZ',
    'Arkansas':'AR',
    'California':'CA',
    'Colorado':'CO',
    'Connecticut':'CT',
    'Delaware':'DE',
    'District of Columbia':'DC',
    'Florida':'FL',
    'Georgia':'GA',
    'Hawaii':'HI',
    'Idaho':'ID',
    'Illinois':'IL',
    'Indiana':'IN',
    'Iowa':'IA',
    'Kansas':'KS',
    'Kentucky':'KY',
    'Louisiana':'LA',
    'Maine':'ME',
    'Maryland':'MD',
    'Massachusetts':'MA',
    'Michigan':'MI',
    'Minnesota':'MN',
    'Mississippi':'MS',
    'Missouri':'MO',
    'Montana':'MT',
    'Nebraska':'NE',
    'Nevada':'NV',
    'New Hampshire':'NH',
    'New Jersey':'NJ',
    'New Mexico':'NM',
    'New York':'NY',
    'North Carolina':'NC',
    'North Dakota':'ND',
    'Ohio':'OH',
    'Oklahoma':'OK',
    'Oregon':'OR',
    'Pennsylvania':'PA',
    'Rhode Island':'RI',
    'South Carolina':'SC',
    'South Dakota':'SD',
    'Tennessee':'TN',
    'Texas':'TX',
    'Utah':'UT',
    'Vermont':'VT',
    'Virginia':'VA',
    'Washington':'WA',
    'West Virginia':'WV',
    'Wisconsin':'WI',
    'Wyoming':'WY',
}


#read data from excel file
file_location = 'C:\git\ElectricityProjectPy\StateElectricityProfiles2007To2016.xlsx'
data2016 = pandas.read_excel(file_location,'DataOf2016')
data2015 = pandas.read_excel(file_location,'DataOf2015')
data2014 = pandas.read_excel(file_location,'DataOf2014')
data2013 = pandas.read_excel(file_location,'DataOf2013')
data2012 = pandas.read_excel(file_location,'DataOf2012')
data2011 = pandas.read_excel(file_location,'DataOf2011')
data2010 = pandas.read_excel(file_location,'DataOf2010')
data2009 = pandas.read_excel(file_location,'DataOf2009')
data2008 = pandas.read_excel(file_location,'DataOf2008')
data2007 = pandas.read_excel(file_location,'DataOf2007')

#todo: Change data file tab to print
print(data2015)
try:
    data2015['Name'] = data2015['Name'].map(us_state)
except KeyError:
    pass


#todo: Change data file tab create certain year report
for col in data2015.columns:
    data2015[col] = data2015[col].astype(str)
scl = [[0.0, 'rgb(49,54,149)'], [0.109999999, 'rgb(69,117,180)'], \
       [0.209999999, 'rgb(116,173,209)'], [0.309999999, 'rgb(171,217,233)'], \
       [0.409999999, 'rgb(224,243,248)'], [0.509999999, 'rgb(254,224,209)'], \
       [0.609999999, 'rgb(253,217,189)'], [0.709999999, 'rgb(244,174,150)'], \
       [0.809999999, 'rgb(215,150,139)'], [0.909999999, 'rgb(180,109,87)'],\
       [1.0,'rgb(175,80,67)']]

#todo: Change data file tab to print
data2015['text'] = data2015['Name']+'<br>'+'Retail Price'+data2015['Average retail price (cents/kWh)']

#todo: Change data file tab to print
data = [dict(
    type='choropleth',
    colorscale = scl,
    autocolorscale = False,
    locations = data2015['Name'],
    z = data2015['Average retail price (cents/kWh)'].astype(float),
    locationmode = 'USA-states',
    text = data2015['text'],
    marker = dict(
        line = dict(
            color = 'rgb(255,255,255)',
            width = 2
                   )),
    colorbar = dict(
        title = "cents/kWh")
             )]


#todo: Change title name to target file
layout = dict(
    title = '2015 State Electricity Profile Average Retail Price<br>(Hover for breakdown)',
    geo = dict(
        scope = 'usa',
        projection = dict(type='albers usa'),
        showlakes = True,
        lakecolor = 'rgb(255,255,255)'),
    )

fig = dict( data=data, layout=layout)
py.iplot(fig, filename='d3-choropleth-map2015AvgRetailPrice')



