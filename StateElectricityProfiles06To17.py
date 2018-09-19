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
print(data2007)
try:
    data2007['Name'] = data2007['Name'].map(us_state)
except KeyError:
    pass


#todo: Change data file tab create certain year report
for col in data2007.columns:
    data2007[col] = data2007[col].astype(str)
scl = [[0.0, 'rgb(49,54,149)'], [0.009999999, 'rgb(69,117,180)'], \
       [0.019999999, 'rgb(116,173,209)'], [0.029999999, 'rgb(171,217,233)'], \
       [0.039999999, 'rgb(224,243,248)'], [0.049999999, 'rgb(254,224,209)'], \
       [0.059999999, 'rgb(253,217,189)'], [0.069999999, 'rgb(244,174,150)'], \
       [0.079999999, 'rgb(215,150,139)'], [0.089999999, 'rgb(180,109,87)'],\
       [0.099999999,'rgb(175,80,67)'],[0.109999999,'rgb(175,70,57)'],\
       [0.119999999,'rgb(170,48,39)'],[1,'rgb(165,0,38)']]

#todo: Change data file tab to print
data2007['text'] = data2007['Name']+'<br>'+'Total Consumption'+data2007['Total retail sales (MWh)']+\
                                '<br>'+'Retail Price'+data2007['Average retail price (cents/kWh)']

#todo: Change data file tab to print
data = [dict(
    type='choropleth',
    colorscale = scl,
    autocolorscale = False,
    locations = data2007['Name'],
    z = data2007['Total retail sales (MWh)'].astype(float),
    locationmode = 'USA-states',
    text = data2007['text'],
    marker = dict(
        line = dict(
            color = 'rgb(255,255,255)',
            width = 2
                   )),
    colorbar = dict(
        title = "MWH")
             )]


#todo: Change title name to target file
layout = dict(
    title = '2007 State Electricity Profile<br>(Hover for breakdown)',
    geo = dict(
        scope = 'usa',
        projection = dict(type='albers usa'),
        showlakes = True,
        lakecolor = 'rgb(255,255,255)'),
    )

fig = dict( data=data, layout=layout)
py.iplot(fig, filename='d3-choropleth-map2007')



