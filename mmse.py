# -*- coding: utf-8 -*-
"""
Created on Wed Jan 21 15:59:38 2015

@author: poggiali
"""
print("Welcome!")
name=raw_input("Patient's name? ")
age=input("Age? ")
name=raw_input("Sex? ")
sch=input("Years of school? ")
PG=input("PG? ")
mat=[[.4,.7,1,1.5,2.2],[-1.1,-.7,-.3,.4,1.4],[-2,-1.6,-1,-.3,.8],[-2.8,-2.3,-1.7,-.9,.3]]

#select column by age
if age>=85:
    j=4
elif age>=80:
    j=3
elif age >=75:
    j=2
elif age>=70:
    j=1    
else:
    j=0

 #select row by school years   
if sch>=13:
    i=3
elif sch>=8:
    i=2
elif sch >=5:
    i=1  
else:
    i=0

PC=PG+mat[i][j]  
if PC>=27:
    val="nella norma"
elif PC>=25:
    val='borderline'
elif PC >=20:
    val='deficitario lieve'
elif PC>=12:
    val='deficit grave'
    j=0

print("Name              PC            Valutazione")
print(name+'   '+str(PC,2)+'    '+val)    
    