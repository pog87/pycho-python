# -*- coding: utf-8 -*-
"""
Created on Mon Feb 23 17:30:57 2015

@author: davide poggiali
@mail poggiali.davide@gmail.com
#some rights reserved; please email me before using/modifying this program
"""

import Tkinter, tkFileDialog

root = Tkinter.Tk()
root.withdraw()

file_path = tkFileDialog.askopenfilename(title="Open original database")


filename=file_path.split('/')[-1].split('.')[0]
extension='.'.join(file_path.split('/')[-1].split('.')[1:])
filepath='/'.join(file_path.split('/')[0:-1])+'/'


import pandas as pd
import matplotlib, numpy as np, scipy
#pd.set_option('display.height', 500)
pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 500)
pd.set_option('max_colwidth',100)


xls = pd.ExcelFile(file_path)
df1 = xls.parse('Foglio1')
df2 = xls.parse('Foglio2')
#df3 = xls.parse('Foglio3')


#functions for RAO
def corretion(raw,mean,b,gen):
    return raw-b[0]*(age-mean[0])-b[1]*(edu-mean[1])-b[2]*gen
def zscore(score,m):
    return (score-m[0])/m[1]

#functions for BICAMS
def scale_cvlt(cvlt_g):
    return 18*(cvlt_g>74)+16*(cvlt_g<=74 and cvlt_g>=72)+15*(cvlt_g<=71 and cvlt_g>=70)+14*(cvlt_g<=69 and cvlt_g>=66)+13*(cvlt_g<=65 and cvlt_g>=63)+12*(cvlt_g<=62 and cvlt_g>=61)+11*(cvlt_g<=60 and cvlt_g>=58)+10*(cvlt_g<=57 and cvlt_g>=55)+9*(cvlt_g<=54 and cvlt_g>=52)+8*(cvlt_g<=51 and cvlt_g>=48)+7*(cvlt_g<=47 and cvlt_g>=45)+6*(cvlt_g<=44 and cvlt_g>=42)+5*(cvlt_g<=41 and cvlt_g>=38)+4*(cvlt_g<=37 and cvlt_g>=35)+2*(cvlt_g<35)
def scale_bvmt(bvmt_g):
    return 18*(bvmt_g>=36)+15*(bvmt_g==35)+14*(bvmt_g==34)+13*(bvmt_g==33)+12*(bvmt_g==32)+11*(bvmt_g==31 or bvmt_g==30)+10*(bvmt_g==28 or bvmt_g==29)+9*(bvmt_g<=27 and bvmt_g>=25)+8*(bvmt_g<=24 and bvmt_g>=23)+7*(bvmt_g<=22 and bvmt_g>=20)+6*(bvmt_g<=19 and bvmt_g>=17)+5*(bvmt_g<=16 and bvmt_g>=14)+4*(bvmt_g<=13 and bvmt_g>=12)+3*(bvmt_g==11)+2*(bvmt_g<11)



#execute for RAO
for i in range(0,df1.shape[0]):
    age=df1.iat[i,3]
    edu=df1.iat[i,5]
    gen=1*(df1.iat[i,3] in ['F','f'])-1*(df1.iat[i,3] in ['M','m'])
    #-1 for male 1 for female
    s=['','.1','.2','.3']
    #cicle for test from 0 to 3
    for j in range(0,4):
        if df1['Forma RAO'+s[j]][i] in ['a', 'A']:
            #age and edu correction terms
            mean=[41.7, 12.4]
            #age, edu, gender correction coefficients
            B_lts=[0,1.402,0]
            B_cltr=[0,1.542,0]
            B_spart=[0,0.368,0]
            B_sdmt=[0,1.029,0]
            B_pasat3=[0,1.698,0]
            B_pasat2=[0,1.116,0]
            B_srtd=[0,0.201,0]
            B_spartD=[0,0.128,0]
            B_wlg=[0,0,2.123]
            B_st=[0.355,0,0.624]
            #mean and standard deviation M
            M_lts=[47.4,14.0]
            M_cltr=[40.2,15.4]
            M_spart=[20.9,5.0]
            M_sdmt=[50.8,10.0]
            M_pasat3=[44.9,12.1]
            M_pasat2=[36.4,12.1]
            M_srtd=[8.9,2.3]
            M_spartD=[7.2,2.4]
            M_wlg=[26.8,5.8]
            M_st=[60.2,12.8]
        elif df1['Forma RAO'+s[j]][i] in ['b', 'B']:
            mean=[43.9,12.7]
            B_lts=[-0.44,0.766,0]
            B_cltr=[-0.457,0.815,0]
            B_spart=[-0.140,0,0]
            B_sdmt=[-0.326,0.871,0]
            B_pasat3=[-0.261,1.177,0]
            B_pasat2=[-0.292,0.935,0]
            B_srtd=[-0.085,0.126,0]
            B_spartD=[-0.054,0,0]
            B_wlg=[0,-0.105,2.904]
            B_st=[0.355,0,0.624]    
            #mean and standard deviation M
            M_lts=[48.1,13.2]
            M_cltr=[39.7,14.9]
            M_spart=[23.1,4.8]
            M_sdmt=[55.3,12.2]
            M_pasat3=[46.3,12.1]
            M_pasat2=[37.6,12.4]
            M_srtd=[9.0,2.4]
            M_spartD=[7.9,1.9]
            M_wlg=[28.5,5.5]
            M_st=[60.2,12.8]  
        #here comes the worst part..
        v=df1['LTS g'+s[j]][i]
        v_corr=corretion(v,mean,B_lts,gen)
        v_z=zscore(v_corr,M_lts)
        df1.ix[i,'LTS c'+s[j]]=v_corr
        df1.ix[i,'LTS z'+s[j]]=v_z
        
        v=df1['CLTR g'+s[j]][i]
        v_corr=corretion(v,mean,B_cltr,gen)
        v_z=zscore(v_corr,M_cltr)
        df1.ix[i,'CLTR c'+s[j]]=v_corr
        df1.ix[i,'CLTR z'+s[j]]=v_z
        
        v=df1['SPART g'+s[j]][i]
        v_corr=corretion(v,mean,B_spart,gen)
        v_z=zscore(v_corr,M_spart)
        df1.ix[i,'SPART c'+s[j]]=v_corr
        df1.ix[i,'SPART z'+s[j]]=v_z
        
        v=df1['SDMT g'+s[j]][i]
        v_corr=corretion(v,mean,B_sdmt,gen)
        v_z=zscore(v_corr,M_sdmt)
        df1.ix[i,'SDMT c'+s[j]]=v_corr
        df1.ix[i,'SDMT z'+s[j]]=v_z
        
        v=df1['PASAT g'+s[j]][i]
        v_corr=corretion(v,mean,B_pasat3,gen)
        v_z=zscore(v_corr,M_pasat3)
        df1.ix[i,'PASAT c'+s[j]]=v_corr
        df1.ix[i,'PASAT z'+s[j]]=v_z
        
        v=df1['SRT-D g'+s[j]][i]
        v_corr=corretion(v,mean,B_srtd,gen)
        v_z=zscore(v_corr,M_srtd)
        df1.ix[i,'SRT-D c'+s[j]]=v_corr
        df1.ix[i,'SRT-D z'+s[j]]=v_z
        
        v=df1['SPART-D g'+s[j]][i]
        v_corr=corretion(v,mean,B_spartD,gen)
        v_z=zscore(v_corr,M_spartD)
        df1.ix[i,'SPART-D c'+s[j]]=v_corr
        df1.ix[i,'SPART-D z'+s[j]]=v_z
        
        v=df1['WLG g'+s[j]][i]
        v_corr=corretion(v,mean,B_wlg,gen)
        v_z=zscore(v_corr,M_wlg)
        df1.ix[i,'WLG c'+s[j]]=v_corr
        df1.ix[i,'WLG z'+s[j]]=v_z
        
        v=df1['STROOP g'+s[j]][i]
        v_corr=corretion(v,mean,B_st,gen)
        v_z=zscore(v_corr,M_st)
        df1.ix[i,'STROOP c'+s[j]]=v_corr
        df1.ix[i,'STROOP z'+s[j]]=v_z
        
        
#check if BICAMS test have been performed and correct the values
#attention! BICAMS scores MUST be an integer!
if 'BICAMS' in df1.columns or 'BICAMS ' in df1.columns:
    #cicle for patient
    for i in range(0,df1.shape[0]):
        age=df1.iat[i,3]
        edu=df1.iat[i,5]
        sex=1+(df1.iat[i,3] in ['F','f'])
        #1 for male 2 for female
        s=['','.1','.2','.3']
        #cicle for test from 0 to 3
        for j in range(0,4):
            cvlt_g=df1['CVLT II g'+s[j]][i]
            bvmt_g=df1['BVMT-R g'+s[j]][i]
            scaled_cvlt=scale_cvlt(cvlt_g)
            scaled_bvmt=scale_cvlt(bvmt_g)
            Ex_value_cvlt=4.989+age*0.118+(age**2)*(-0.002)+sex*0.823+edu*0.178
            Ex_value_bvmt=12.694+age*(-0.039)+(age**2)*(-0.001)+sex*(-0.671)+edu*(0.101)
            z_cvlt=(Ex_value_cvlt-scaled_cvlt)/(2.842)
            z_bvmt=(Ex_value_bvmt-scaled_bvmt)/(3.133)
            t_cvlt=50+10*z_cvlt
            t_bvmt=50+10*z_bvmt
            df1.ix[i,'CVLT II z'+s[j]]=z_cvlt
            df1.ix[i,'CVLT II t'+s[j]]=t_cvlt
            df1.ix[i,'BVMT-R z'+s[j]]=z_bvmt
            df1.ix[i,'BVMT-R t'+s[j]]=t_bvmt
#save whole databese
file2=filepath+filename+'_mod.'+extension        
writer = pd.ExcelWriter(file2)
df1.to_excel(writer, 'Foglio1')
df2.to_excel(writer, 'Foglio2')
writer.save()