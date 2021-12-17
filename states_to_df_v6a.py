# -*- coding: utf-8 -*-
"""
v4 - save excel as pdf, format excel nicely
v3h - import gmail PY file
v3d - change directory, moved "COVID_TheAtlantic" up one level
v3a - add USA as a whole
@author: Kyle
"""
#%%
#from xlwt import Workbook
#from SubDirTestA_forV3.SubDirTestB.betas_samp_dist import betas_samp
import requests
#import pypiwin32 # FOR XLS->PDF
import win32com.client # FOR XLS->PDF
#from PIL import Image
#import numpy as np
from datetime import datetime
now = datetime.now() 
dtime_string = now.strftime("%Y-%m-%d---%H-%M-%S")
import pandas as pd
import os
from pandas import DataFrame
covid_path = 'C:\\Users\\Kyle\\Desktop\\_FOLDERS_\\COVID_TheAtlantic\\Historic state data\\'
covid_top_path = 'C:\\Users\\Kyle\\Desktop\\_FOLDERS_\\COVID_TheAtlantic\\'
states_path = 'C:\\Users\\Kyle\\Desktop\\_FOLDERS_\\COVID_TheAtlantic\\Historic state data\\daily.csv'

#%%
#df_old_old = pd.read_csv(states_path)
#dir_str = str('C:/Users/Kyle/Desktop/PYTHON/OPTIONS_DLs/Yahoo/')
ss55 = now
#DateDirStr = str(dir_str)+str('_A_/')
DateDirStr = str(covid_path)+str('HISTORY\\')+str(ss55)[0:10]+\
    str('_')+str(ss55)[11:13]+str('h')+str(ss55)[14:16]+str('m')+str(ss55)[17:19]+str('s/')+str('\\')
os.mkdir(DateDirStr)
url_states = 'https://covidtracking.com/api/v1/states/daily.csv'
r = requests.get(url_states, allow_redirects=True)
states_csv_newName = DateDirStr + str('daily_CSV_') + dtime_string + str('.csv')
open(states_csv_newName, 'wb').write(r.content)

#%%
df = pd.read_csv(states_csv_newName)

# %%
# Grab "daily.csv" file for USA totals (not broken out by state)
UsaDateDirStr = str(covid_top_path)+str('Historic US data\\HISTORY\\')+str(ss55)[0:10]+\
    str('_')+str(ss55)[11:13]+str('h')+str(ss55)[14:16]+str('m')+str(ss55)[17:19]+str('s/')
os.mkdir(UsaDateDirStr)
url_USA = 'https://covidtracking.com/api/v1/us/daily.csv'
r = requests.get(url_USA, allow_redirects=True)
usa_csv_newName = UsaDateDirStr + str('daily_CSV_') + dtime_string + str('.csv')
open(usa_csv_newName, 'wb').write(r.content)

# %%
class Error(Exception):
    """Base class for exceptions in this module."""
    # "Error class inherits from Exception class":
    # https://docs.python.org/3/tutorial/errors.html#user-defined-exceptions
    pass
class InputError(Error):
    """Exception raised for errors in the input.

    Attributes:
        expression -- input expression in which the error occurred
        message -- explanation of the error
    """
#    def __init__(self, expression, message):
#        self.expression = expression
#        self.message = message
def __init__(self, message):
    self.message = message

# %%
# Initialize size of new dataframe:
n_dates = df['date'].nunique()
n_states = df['state'].nunique()

# First use of numpy:
u_dates = list(set(df['date']))
u_states = list(set(df['state']))

u_dates.sort() # ooof this sorts the dates as if they're STRINGS, not what we want...
u_states.sort()

datetimesss = pd.to_datetime(u_dates,format='%Y%m%d') # yay this "parses"
ddd = datetimesss.sort_values()
# self, by, axis=0, ascending=True, inplace=False
ddd2 = datetimesss.sort_values(ascending=False)
ddd=ddd2.copy()
del datetimesss
del ddd2

# Initialize DF:
# yay, initializes to nans
cum_cases = DataFrame(index=ddd, columns=u_states)
cum_deaths = DataFrame(index=ddd, columns=u_states)
new_cases = DataFrame(index=ddd, columns=u_states)
new_deaths = DataFrame(index=ddd, columns=u_states)
totalTestResults = DataFrame(index=ddd, columns=u_states)
hospitalizedCurrently = DataFrame(index=ddd, columns=u_states)
inIcuCurrently = DataFrame(index=ddd, columns=u_states)
onVentilatorCurrently = DataFrame(index=ddd, columns=u_states)

# Now populate with every entry from "df" (the raw rectangle format csv):
for ii in range(0,df.shape[0]):
    # Find date and state:
    state_ii = df['state'][ii]
    date_ii_str = df['date'][ii]
    date_ii = pd.to_datetime(date_ii_str,format='%Y%m%d') 
    cum_cases_ii = df['positive'][ii]
    cum_deaths_ii = df['death'][ii]
    new_cases_ii = df['positiveIncrease'][ii]
    new_deaths_ii = df['deathIncrease'][ii]
    totalTestResults_ii = df['totalTestResults'][ii]
    hospitalizedCurrently_ii = df['hospitalizedCurrently'][ii]
    inIcuCurrently_ii = df['inIcuCurrently'][ii]
    onVentilatorCurrently_ii = df['onVentilatorCurrently'][ii]
    # Now populate:
    cum_cases.loc[date_ii][state_ii] = cum_cases_ii
    cum_deaths.loc[date_ii][state_ii] = cum_deaths_ii
    new_cases.loc[date_ii][state_ii] = new_cases_ii
    new_deaths.loc[date_ii][state_ii] = new_deaths_ii
    totalTestResults.loc[date_ii][state_ii] = totalTestResults_ii
    hospitalizedCurrently.loc[date_ii][state_ii] = hospitalizedCurrently_ii
    inIcuCurrently.loc[date_ii][state_ii] = inIcuCurrently_ii
    onVentilatorCurrently.loc[date_ii][state_ii] = onVentilatorCurrently_ii

    
del state_ii
del date_ii
del cum_cases_ii
del cum_deaths_ii
del totalTestResults_ii
del new_cases_ii
del new_deaths_ii
del ii

# %% 
# First, create "Past 7 days avg" Dataframes for new_cases, new_deaths:
new_cases_7d = DataFrame(index=ddd, columns=u_states)
new_deaths_7d = DataFrame(index=ddd, columns=u_states)
for rr in range(0,new_deaths.shape[0]-7):
    for cc in range(0,new_deaths.shape[1]):
#        avg_week_cases_ii = sum(new_cases_7d.iloc[rr-6:rr][cc])/7
#        avg_week_deaths_ii = sum(new_deaths_7d.iloc[rr-6:rr][cc])/7
#        new_cases_7d.iloc[rr][cc] = avg_week_cases_ii
#        new_deaths_7d.iloc[rr][cc] = avg_week_deaths_ii
        aa = 0
        bb = 0
        for ii in range(rr,rr+7):
            aa = aa + new_cases.iloc[ii][cc]
            bb = bb + new_deaths.iloc[ii][cc]
        new_cases_7d.iloc[rr][cc] = aa/7
        new_deaths_7d.iloc[rr][cc] = bb/7
        del aa
        del bb
        del ii

# %%
stats_01 = DataFrame(index=u_states, columns=pd.RangeIndex(1,13))
aaa=cum_cases.iloc[0][:].rank(ascending=False)
bbb=cum_deaths.iloc[0][:].rank(ascending=False)
ccc=new_cases.iloc[0][:].rank(ascending=False)
ddd=new_deaths.iloc[0][:].rank(ascending=False)
eee=new_cases_7d.iloc[0][:].rank(ascending=False)
fff=new_deaths_7d.iloc[0][:].rank(ascending=False)
for ss in u_states:
    # (1a) Rank total cases (by state)
    # (1b) Rank total deaths (by state)
    # (1c) Rank NEW cases
    # (1d) Rank NEW deaths
    # (1e) Rank NEW cases [7-day]
    # (1f) Rank NEW deaths [7-day]
    stats_01.loc[ss][1]=cum_cases.iloc[0][ss]
    stats_01.loc[ss][2]=aaa[ss]
    stats_01.loc[ss][3]=cum_deaths.iloc[0][ss]
    stats_01.loc[ss][4]=bbb[ss]
    stats_01.loc[ss][5]=new_cases.iloc[0][ss]
    stats_01.loc[ss][6]=ccc[ss]
    stats_01.loc[ss][7]=new_deaths.iloc[0][ss]
    stats_01.loc[ss][8]=ddd[ss]
    stats_01.loc[ss][9]=new_cases_7d.iloc[0][ss]
    stats_01.loc[ss][10]=eee[ss]
    stats_01.loc[ss][11]=new_deaths_7d.iloc[0][ss]
    stats_01.loc[ss][12]=fff[ss]
#del aaa
#del bbb
#del ccc
#del ddd
#del eee
#del fff
# Redfine column names:
stats_01.columns = ['CumCases','CumCases_RANK','CumDeaths','CumDeaths_RANK',\
                    'NewCases','NewCases_RANK','NewDeaths','NewDeaths_RANK',\
                    'NewCases_7d','NewCases_7d_RANK','NewDeaths_7d','NewDeaths_7d_RANK']


# %% v0f: add images before data sheets:
#
# https://stackoverflow.com/questions/45376232/how-to-save-image-created-with-pandas-dataframe-plot/45379210
#
plotA = new_cases.plot(y=['MI','PA','SC','NY'],color=['red','green',\
               'blue','yellow'])
plotA.set_ylabel("New COVID Cases Per Day")
figA=plotA.get_figure()
figAstr = DateDirStr + str('NewCASES_imgD_') + str(dtime_string) + str('.png')
figA.savefig(figAstr)
del figA
del plotA
plotB = new_cases.plot(y=['IL','NJ','CA'],color=['red',\
               'blue','black'])
plotB.set_ylabel("New COVID Cases Per Day")
figB=plotB.get_figure()
figBstr = DateDirStr + str('NewCASES_imgB_') + str(dtime_string) + str('.png')
figB.savefig(figBstr)
del figB
del plotB
plotB = new_cases.plot(y=['MA','FL','TX'],color=['red',\
               'blue','black'])
plotB.set_ylabel("New COVID Cases Per Day")
figB=plotB.get_figure()
figBstr = DateDirStr + str('NewCASES_imgC_') + str(dtime_string) + str('.png')
figB.savefig(figBstr)
del figB
del plotB
plotB = new_cases.plot(y=['PA','MI','SC'],color=['red',\
               'blue','black'])
plotB.set_ylabel("New COVID Cases Per Day")
figB=plotB.get_figure()
figBstr = DateDirStr + str('NewCASES_imgA_') + str(dtime_string) + str('.png')
figB.savefig(figBstr)
del figB
del plotB
plotB = new_cases.plot(y=['AK','AL','GA','NC'],color=['red','green',\
               'blue','black'])
plotB.set_ylabel("New COVID Cases Per Day")
figB=plotB.get_figure()
figBstr = DateDirStr + str('NewCASES_imgE_') + str(dtime_string) + str('.png')
figB.savefig(figBstr)
del figB
del plotB

######################################################
# AVERAGE (7-DAY) IMAGES - NEW CASES [NEW DEATHS] BELOW
######################################################

plotA = new_cases_7d.plot(y=['MI','PA','SC','NY'],color=['red','green',\
               'blue','yellow'])
plotA.set_ylabel("7D AVG - New COVID Cases Per Day")
figA=plotA.get_figure()
figAstr = DateDirStr + str('AVG_Cases_imgD_') + str(dtime_string) + str('.png')
figA.savefig(figAstr)
del figA
del plotA
plotB = new_cases_7d.plot(y=['IL','NJ','CA'],color=['red',\
               'blue','black'])
plotB.set_ylabel("7D AVG - New COVID Cases Per Day")
figB=plotB.get_figure()
figBstr = DateDirStr + str('AVG_Cases_imgB_') + str(dtime_string) + str('.png')
figB.savefig(figBstr)
del figB
del plotB
plotB = new_cases_7d.plot(y=['MA','FL','TX'],color=['red',\
               'blue','black'])
plotB.set_ylabel("7D AVG - New COVID Cases Per Day")
figB=plotB.get_figure()
figBstr = DateDirStr + str('AVG_Cases_imgC_') + str(dtime_string) + str('.png')
figB.savefig(figBstr)
del figB
del plotB
plotB = new_cases_7d.plot(y=['PA','MI','SC'],color=['red',\
               'blue','black'])
plotB.set_ylabel("7D AVG - New COVID Cases Per Day")
figB=plotB.get_figure()
figBstr = DateDirStr + str('AVG_Cases_imgA_') + str(dtime_string) + str('.png')
figB.savefig(figBstr)
del figB
del plotB
plotB = new_cases_7d.plot(y=['AK','AL','GA','NC'],color=['red','green',\
               'blue','black'])
plotB.set_ylabel("7D AVG - New COVID Cases Per Day")
figB=plotB.get_figure()
figBstr = DateDirStr + str('AVG_Cases_imgE_') + str(dtime_string) + str('.png')
figB.savefig(figBstr)
del figB
del plotB

######################################################
# AVERAGE (7-DAY) IMAGES - NEW **DEATHS** deaths deaths **DEATHS**
######################################################

plotA = new_deaths_7d.plot(y=['MI','PA','SC','NY'],color=['red','green',\
               'blue','yellow'])
plotA.set_ylabel("7D AVG - New COVID DEATHS Per Day")
figA=plotA.get_figure()
figAstr = DateDirStr + str('DEATHS_imgD_') + str(dtime_string) + str('.png')
figA.savefig(figAstr)
del figA
del plotA
plotB = new_deaths_7d.plot(y=['IL','NJ','CA'],color=['red',\
               'blue','black'])
plotB.set_ylabel("7D AVG - New COVID DEATHS Per Day")
figB=plotB.get_figure()
figBstr = DateDirStr + str('DEATHS_imgB_') + str(dtime_string) + str('.png')
figB.savefig(figBstr)
del figB
del plotB
plotB = new_deaths_7d.plot(y=['MA','FL','TX'],color=['red',\
               'blue','black'])
plotB.set_ylabel("7D AVG - New COVID DEATHS Per Day")
figB=plotB.get_figure()
figBstr = DateDirStr + str('DEATHS_imgC_') + str(dtime_string) + str('.png')
figB.savefig(figBstr)
del figB
del plotB
plotB = new_deaths_7d.plot(y=['PA','MI','SC'],color=['red',\
               'blue','black'])
plotB.set_ylabel("7D AVG - New COVID DEATHS Per Day")
figB=plotB.get_figure()
figBstr = DateDirStr + str('DEATHS_imgA_') + str(dtime_string) + str('.png')
figB.savefig(figBstr)
del figB
del plotB
plotB = new_deaths_7d.plot(y=['AK','AL','GA','NC'],color=['red','green',\
               'blue','black'])
plotB.set_ylabel("7D AVG - New COVID DEATHS Per Day")
figB=plotB.get_figure()
figBstr = DateDirStr + str('DEATHS_imgE_') + str(dtime_string) + str('.png')
figB.savefig(figBstr)
del figB
del plotB

######################################################
######################################################
######################################################

# %%
#stats_01.to_excel(write_path,sheet_name='Stats01')
# (2) CLOSER LOOK AT FOUR STATES ('MI','PA','SC','NY')
# (2a) total cases (& rank)
# (2b) total deaths (& rank)
# (2c) NEW cases (& rank)
# (2d) NEW deaths (& rank)
# (2e) NEW cases [7-day] (& rank)
# (2f) NEW deaths [7-day] (& rank)
# (2g1-4) PEAK/MAX for (2c-2g) [& date]
# (2h1-4) % change of (2c-2g) relative to MAX
# %%
stats_02 = DataFrame(index=['NY','PA','SC','MI','IL','NJ','FL','GA','NC','MA','TX','MD'], columns=pd.RangeIndex(1,41))
for ss in ['NY','PA','SC','MI','IL','NJ','FL','GA','NC','MA','TX','MD']:
#stats_02 = DataFrame(index=u_states, columns=pd.RangeIndex(1,41))
#for ss in u_states:
    stats_02.loc[ss][1]=cum_cases.iloc[0][ss]
    stats_02.loc[ss][2]=aaa[ss]
    stats_02.loc[ss][3]=cum_deaths.iloc[0][ss]
    stats_02.loc[ss][4]=bbb[ss]
    stats_02.loc[ss][5]=new_cases.iloc[0][ss]
    stats_02.loc[ss][6]=ccc[ss]
    stats_02.loc[ss][7]=new_deaths.iloc[0][ss]
    stats_02.loc[ss][8]=ddd[ss]
    stats_02.loc[ss][9]=new_cases_7d.iloc[0][ss]
    stats_02.loc[ss][10]=eee[ss]
    stats_02.loc[ss][11]=new_deaths_7d.iloc[0][ss]
    stats_02.loc[ss][12]=fff[ss]
    # Find max new cases:
    stats_02.loc[ss][13]=max(new_cases[ss])
#    zz=new_cases[ss]
    stats_02.loc[ss][14]=max(new_cases[ss].index, key=new_cases[ss].get)
    stats_02.loc[ss][15]=max(new_deaths[ss])
#    zz=new_deaths[ss]
    stats_02.loc[ss][16]=max(new_deaths[ss].index, key=new_deaths[ss].get)
    stats_02.loc[ss][17]=max(new_cases_7d[ss])
#    zz=new_cases_7d[ss]
    stats_02.loc[ss][18]=max(new_cases_7d[ss].index, key=new_cases_7d[ss].get)
    stats_02.loc[ss][19]=max(new_deaths_7d[ss])
#    zz=new_deaths_7d[ss]
    stats_02.loc[ss][20]=max(new_deaths_7d[ss].index, key=new_deaths_7d[ss].get)
    stats_02.loc[ss][21]=(new_cases.iloc[0][ss] - max(new_cases[ss]))/max(new_cases[ss])
    stats_02.loc[ss][22]=(new_deaths.iloc[0][ss] - max(new_deaths[ss]))/max(new_deaths[ss])
    stats_02.loc[ss][23]=(new_cases_7d.iloc[0][ss] - max(new_cases_7d[ss]))/max(new_cases_7d[ss])
    stats_02.loc[ss][24]=(new_deaths_7d.iloc[0][ss] - max(new_deaths_7d[ss]))/max(new_deaths_7d[ss])
    # v1e - more weekly change stats:
    stats_02.loc[ss][25]=new_cases_7d.iloc[6][ss]
    stats_02.loc[ss][26]=(new_cases_7d.iloc[0][ss]- new_cases_7d.iloc[6][ss])/new_cases_7d.iloc[6][ss]
    stats_02.loc[ss][27]=new_deaths_7d.iloc[6][ss]
    stats_02.loc[ss][28]=(new_deaths_7d.iloc[0][ss]- new_deaths_7d.iloc[6][ss])/new_deaths_7d.iloc[6][ss]
    stats_02.loc[ss][29]=new_cases_7d.iloc[13][ss]
    stats_02.loc[ss][30]=(new_cases_7d.iloc[0][ss]- new_cases_7d.iloc[13][ss])/new_cases_7d.iloc[13][ss]
    stats_02.loc[ss][31]=new_deaths_7d.iloc[13][ss]
    stats_02.loc[ss][32]=(new_deaths_7d.iloc[0][ss]- new_deaths_7d.iloc[13][ss])/new_deaths_7d.iloc[13][ss]
    stats_02.loc[ss][33]=new_cases_7d.iloc[20][ss]
    stats_02.loc[ss][34]=(new_cases_7d.iloc[0][ss]- new_cases_7d.iloc[20][ss])/new_cases_7d.iloc[20][ss]
    stats_02.loc[ss][35]=new_deaths_7d.iloc[20][ss]
    stats_02.loc[ss][36]=(new_deaths_7d.iloc[0][ss]- new_deaths_7d.iloc[20][ss])/new_deaths_7d.iloc[20][ss]
    stats_02.loc[ss][37]=new_cases_7d.iloc[27][ss]
    stats_02.loc[ss][38]=(new_cases_7d.iloc[0][ss]- new_cases_7d.iloc[27][ss])/new_cases_7d.iloc[27][ss]
    stats_02.loc[ss][39]=new_deaths_7d.iloc[27][ss]
    stats_02.loc[ss][40]=(new_deaths_7d.iloc[0][ss]- new_deaths_7d.iloc[27][ss])/new_deaths_7d.iloc[27][ss]

#%%
stats_02.columns = ['CumCases','CumCases_RANK','CumDeaths','CumDeaths_RANK',\
                    'NewCases','NewCases_RANK','NewDeaths','NewDeaths_RANK',\
                    'NewCases_7d','NewCases_7d_RANK','NewDeaths_7d','NewDeaths_7d_RANK',\
                    'MAX_NewCases','Max_NewCases_Date','Max_NewDeaths','Max_NewDeaths_Date',\
                    'MAX_NewCases_7d','Max_NewCases_7d_Date','Max_NewDeaths_7d','Max_NewDeaths_7d_Date',\
                    '%Chg_MaxCases','%Chg_MaxDeaths','%Chg_MaxCases_7d','%Chg_MaxDeaths_7d',\
                    'NewCases_7d_LastWk','%Chg_CasesLastWk','NewDeaths_7d_LastWk','%Chg_DeathsLastWk',\
                    'NewCases_7d_2WksAgo','%Chg_Cases_2WksAgo','NewDeaths_7d_2WksAgo','%Chg_Deaths_2WksAgo',\
                    'NewCases_7d_3WksAgo','%Chg_Cases_3WksAgo','NewDeaths_7d_3WksAgo','%Chg_Deaths_3WksAgo',\
                    'NewCases_7d_4WksAgo','%Chg_Cases_4WksAgo','NewDeaths_7d_4WksAgo','%Chg_Deaths_4WksAgo']
stats2trans = stats_02.transpose()
week_over_1week = DataFrame(index=u_states, columns=pd.RangeIndex(1,13))


#stats2trans.to_excel(write_path,sheet_name='Stats02')
#new_cases.plot(y=['MI','PA','SC','NY'],color=['red','green',\
#               'blue','yellow']).to_excel(write_path,sheet_name='Stats03')
# CAN WE OUTPUT PLOTS TO EXCEL...
# %%
#fig = class_counts.plot(kind='bar',  figsize=(20, 16), fontsize=26).get_figure()


# %% ALL THE EXCEL OUTPUT GOES HERE:
write_path = DateDirStr + str('STATS_') + str(dtime_string) + str('.xlsx')
emptyDF = DataFrame()
emptyDF.to_excel(write_path,sheet_name='Empty')
#stats_01.to_excel(write_path,sheet_name='Stats0A')
with pd.ExcelWriter(write_path,
                    engine="openpyxl",
                    mode='a') as writer:  
    stats_01.to_excel(writer,sheet_name='Stats_01')
    stats2trans.to_excel(writer,sheet_name='Max_Stats')
    # v2d - 4 dataframes to excel:
    cum_cases.to_excel(writer,sheet_name='TotalCases')
    cum_deaths.to_excel(writer,sheet_name='TotalDeaths') # looks like this should be deaths
    new_cases.to_excel(writer,sheet_name='NewCases')
    new_deaths.to_excel(writer,sheet_name='NewDeaths')
    new_cases_7d.to_excel(writer,sheet_name='NewCases_7d')
    new_deaths_7d.to_excel(writer,sheet_name='NewDeaths_7d')
    writer.save()

# dataframes to HTML:
html_dir_01 = DateDirStr + 'cum_cases.html'
html_dir_02 = DateDirStr + 'cum_deaths.html'
html_dir_03 = DateDirStr + 'new_cases_7d.html'
html_dir_04 = DateDirStr + 'new_deaths_7d.html'
cum_cases.to_html(html_dir_01)
cum_deaths.to_html(html_dir_02)
new_cases_7d.to_html(html_dir_03)
new_deaths_7d.to_html(html_dir_04)


#    df.to_excel(writer, sheet_name='Sheet_name_3')
###############################################################
# %% ADD USA (and "regions" e.g. NJ+PA) AS A WHOLE:
# use "x" suffix to denote copies of variables, we'll add "USA" to these:
stats_01x = stats_01.copy()
stats_02x = stats_02.copy()
new_cases_x = new_cases.copy()
new_cases_7d_x = new_cases_7d.copy()
new_deaths_x = new_deaths.copy()
new_deaths_7d_x = new_deaths_7d.copy()

new_columns = ['USA','USA_48','PA+NJ','PA+NJ_v2']
new_cases_x.loc[:,new_columns[0]]=(new_cases_x.iloc[:, 0:57]).sum(axis=1)
new_cases_x.loc[:,new_columns[1]]=(new_cases_x.iloc[:,0:1].sum(axis=1))
new_cases_x.loc[:,new_columns[2]]=new_cases_x['NJ'] + new_cases_x['PA']
new_cases_x.loc[:,new_columns[3]]=(new_cases_x.iloc[:, [34,41]]).sum(axis=1)
# New deaths:
new_deaths_x.loc[:,new_columns[0]]=(new_deaths_x.iloc[:, 0:57]).sum(axis=1)
new_deaths_x.loc[:,new_columns[1]]=(new_deaths_x.iloc[:,0:1].sum(axis=1))
new_deaths_x.loc[:,new_columns[2]]=new_deaths_x['NJ'] + new_deaths_x['PA']
new_deaths_x.loc[:,new_columns[3]]=(new_deaths_x.iloc[:, [34,41]]).sum(axis=1)
# Averages:
new_deaths_7d_x.loc[:,new_columns[0]]=(new_deaths_7d_x.iloc[:, 0:57]).sum(axis=1)
new_deaths_7d_x.loc[:,new_columns[1]]=(new_deaths_7d_x.iloc[:,0:1].sum(axis=1))
new_deaths_7d_x.loc[:,new_columns[2]]=new_deaths_7d_x['NJ'] + new_deaths_7d_x['PA']
new_deaths_7d_x.loc[:,new_columns[3]]=(new_deaths_7d_x.iloc[:, [34,41]]).sum(axis=1)
new_cases_7d_x.loc[:,new_columns[0]]=(new_cases_7d_x.iloc[:, 0:57]).sum(axis=1)
new_cases_7d_x.loc[:,new_columns[1]]=(new_cases_7d_x.iloc[:,0:1].sum(axis=1))
new_cases_7d_x.loc[:,new_columns[2]]=new_cases_7d_x['NJ'] + new_cases_7d_x['PA']
new_cases_7d_x.loc[:,new_columns[3]]=(new_cases_7d_x.iloc[:, [34,41]]).sum(axis=1)

##########################################
# %%
# USA PLOTS:
##########################################
plotB = new_cases_7d_x.plot(y=['USA'],color=['red'])
plotB.set_ylabel("7D AVG - New COVID Cases Per Day")
figB=plotB.get_figure()
figBstr = DateDirStr + str('USA_CASES_7dAvg_') + str(dtime_string) + str('.png')
figB.savefig(figBstr)
del figB
del plotB
plotA = new_deaths_7d_x.plot(y=['USA'],color=['black'])
plotA.set_ylabel("7D AVG - New COVID DEATHS Per Day")
figA=plotA.get_figure()
figAstr = DateDirStr + str('USA_DEATHS_7dAvg_') + str(dtime_string) + str('.png')
figA.savefig(figAstr)
del figA
del plotA
plotB = new_cases_x.plot(y=['USA'],color=['red'])
plotB.set_ylabel("New COVID Cases Per Day")
figB=plotB.get_figure()
figBstr = DateDirStr + str('USA_NewCASES_') + str(dtime_string) + str('.png')
figB.savefig(figBstr)
del figB
del plotB
plotB = new_deaths_x.plot(y=['USA'],color=['black'])
plotB.set_ylabel("New COVID DEATHS Per Day")
figB=plotB.get_figure()
figBstr = DateDirStr + str('USA_NewDEATHS_') + str(dtime_string) + str('.png')
figB.savefig(figBstr)
del figB
del plotB


###############################################################
# %% Excel formatting:
###############################################################

##o = win32com.client.Dispatch("Excel.Application")
##o.Visible = False
#
#wb_path = write_path
#format_wb_path = 'C:\\Users\\Kyle\\Desktop\\_FOLDERS_\\COVID_TheAtlantic\\Historic state data\\STATS_FormatPASTE_older.xlsx'
##wb_to = o.Workbooks.Open(wb_path)
##wb_from = o.Workbooks.Open(format_wb_path)
##ws_to = wb_to.Worksheets['Max_Stats']
##ws_from = wb_from.Worksheets['Max_Stats']
#
#import openpyxl as xl
#path1 = format_wb_path
#path2 = wb_path
#
#wb1 = xl.load_workbook(filename=path1)
#ws1 = wb1.worksheets[0]
#
#wb2 = xl.load_workbook(filename=path2)
#ws2 = wb2.worksheets[0]
#
#for row in ws1:
#    for cell in row:
#        ws2[cell.coordinate].value = cell.value
#
#wb2.save(path2)


##########################################
# %%
# Save Excel as PDF:
##########################################

#o = win32com.client.Dispatch("Excel.Application")
#o.Visible = False
#
#wb_path = write_path
#wb = o.Workbooks.Open(wb_path)
#
#ws_index_list = [2] #say you want to print these sheets
#path_to_pdf = DateDirStr + str('STATS_') + str(dtime_string) + str('.pdf')
#
#for index in ws_index_list:
#    ws = wb.Worksheets[index-1]
#    ws.PageSetup.Zoom = False
#    ws.PageSetup.FitToPagesTall = 1
#    ws.PageSetup.FitToPagesWide = 1
#    if index == 2:
#        # Sheet "Stats_01"
#        print_area = 'A1:M57'
#    elif index == 3:
#        # Sheet "Max_Stats"
#        print_area = 'A1:M41'
#    else:
#        break
#        
#    ws.PageSetup.PrintArea = print_area
#
#wb.WorkSheets(ws_index_list).Select()
#wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)
##wb.save() # need this otherwise program gets stuck in background on windows prompt "Do you want to save"
##wb.close()

##########################################

#%%
# EMAIL:

import smtplib
import os
from email.mime.text import MIMEText
#from email.MIMEMultipart import MIMEMultipart
#from email.MIMEBase import MIMEBase
#from email import Encoders
#from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
def send_mail_gmail(username,password,toaddrs_list,\
                    msg_text,fromaddr=None,subject="Test mail",\
                    attachment_path_list=None):

    s = smtplib.SMTP('smtp.gmail.com:587')
    s.starttls()
    s.login(username, password)
    #s.set_debuglevel(1)
    msg = MIMEMultipart()
    sender = fromaddr
    recipients = toaddrs_list
    msg['Subject'] = subject
    if fromaddr is not None:
        msg['From'] = sender
    msg['To'] = ", ".join(recipients)
    if attachment_path_list is not None:
        os.chdir(DateDirStr)
        files = os.listdir()
        for f in files:  # add files to the message`
            try:
                file_path = os.path.join(attachment_path_list, f)
                attachment = MIMEApplication(open(file_path, "rb").read(), _subtype="txt")
                attachment.add_header('Content-Disposition','attachment', filename=f)
                msg.attach(attachment)
            except:
                print("could not attach file")
    msg.attach(MIMEText(msg_text,'html'))
    s.sendmail(sender, recipients, msg.as_string())

subj_str = subject=str('COVID HISTORY ')+str(ss55)[0:10]+\
                str(' (')+str(ss55)[11:13]+str('h')+str(ss55)[14:16]+str('m)')

most_recent_str = str(new_cases.index[0])[0:10]
msg_text_01=str('SOURCES: '+str('\n')+\
             'https://covidtracking.com/api'+str('\n')+\
             'https://www.worldometers.info/coronavirus/'+str('\n')+\
             'https://www.worldometers.info/coronavirus/country/us/'+str('\n')+\
             ' '+str('\n')+\
             'Most recent data point: '+most_recent_str+str('\n'))


stats_text_A =  str('----- --------------------------------- TOP LINE STATS: --------------------------------- -----')+\
                str('USA 7d NEW CASES: ') + str(round(new_cases_7d_x.iloc[0]['USA'],1)) +\
                str('\n')+str('\n')+\
                str('USA 7d NEW DEATHS: ') + str(round(new_deaths_7d_x.iloc[0]['USA'],1)) +\
                str('\n')+str('\n')+\
                str('MICHIGAN 7d NEW CASES: ') + str(round(new_cases_7d_x.iloc[0]['MI'],1)) +\
                str('\n')+str('\n')+\
                str('MICHIGAN 7d NEW DEATHS: ') + str(round(new_deaths_7d_x.iloc[0]['MI'],1)) +\
                str('\n')+str('\n')+\
                str('----- --------------------------------- PREVIOUS WEEK STATS: --------------------------------- -----')+\
                str('\n')+str('\n')+\
                str('USA 7d NEW CASES: ') + str(round(new_cases_7d_x.iloc[6]['USA'],1)) +\
                str('\n')+str('\n')+\
                str('USA 7d NEW DEATHS: ') + str(round(new_deaths_7d_x.iloc[6]['USA'],1)) +\
                str('\n')+str('\n')+\
                str('MICHIGAN 7d NEW CASES: ') + str(round(new_cases_7d_x.iloc[6]['MI'],1)) +\
                str('\n')+str('\n')+\
                str('MICHIGAN 7d NEW DEATHS: ') + str(round(new_deaths_7d_x.iloc[6]['MI'],1))+\
                str('----- --------------------------------- PREVIOUS MONTH STATS: --------------------------------- -----')+\
                str('\n')+str('\n')+\
                str('USA 7d NEW CASES: ') + str(round(new_cases_7d_x.iloc[30]['USA'],1)) +\
                str('\n')+str('\n')+\
                str('USA 7d NEW DEATHS: ') + str(round(new_deaths_7d_x.iloc[30]['USA'],1)) +\
                str('\n')+str('\n')+\
                str('MICHIGAN 7d NEW CASES: ') + str(round(new_cases_7d_x.iloc[30]['MI'],1)) +\
                str('\n')+str('\n')+\
                str('MICHIGAN 7d NEW DEATHS: ') + str(round(new_deaths_7d_x.iloc[30]['MI'],1)) +\
                str('----- ---------------------------------')



msg_text_02=str('SOURCES: '+str('\n')+\
                ' '+str('\n')+\
                'https://covidtracking.com/api'+str('\n')+\
                ' '+str('\n')+\
                'https://www.worldometers.info/coronavirus/'+str('\n')+\
                ' '+str('\n')+\
                'https://www.worldometers.info/coronavirus/country/us/'+str('\n')+\
                ' '+str('\n')+\
                ' '+str('\n')+\
                'Most recent data point: '+most_recent_str+str('\n'))
#msg_text_01=['SOURCES: '+str('\n')+\
#             'https://covidtracking.com/api'+str('\n')+\
#             'https://www.worldometers.info/coronavirus/'+str('\n')+\
#             'https://www.worldometers.info/coronavirus/country/us/'+str('\n')+\
#             ' '+str('\n')+\
#             'Most recent data point: '+most_recent_str+str('\n')]

msg_text_02_w_stats = stats_text_A + str('\n')+\
                ' '+str('\n')+\
                ' '+str('\n') + msg_text_02
                
send_mail_gmail(username='kcb.sender01@gmail.com',password='hi_github_user_-_you_should_encrypt_this',\
                toaddrs_list=['kylebinder14@gmail.com',\
                              'tothereadinglist102@gmail.com'],\
                msg_text=msg_text_02_w_stats,fromaddr='kcb.sender22@gmail.com',\
                subject=subj_str,\
                attachment_path_list=DateDirStr)    


