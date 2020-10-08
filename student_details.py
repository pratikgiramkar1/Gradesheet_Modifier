import pandas as pd
import math
#low=[]
#high=[]
#var=0
df=pd.read_excel('std4.xlsx')
#with pd.option_context('display.max_rows', None, 'display.max_columns', None):  # more options can be specified also
    #print(df)
#for i in range(len(df)) :
  #print(df.loc[i, "Sess1"], df.loc[i, "Subjects"])
#print(df.)
math=[]
sum_math=0
i=0
j=0
while i<len(df) :
  math.append(df.iloc[i, 8])
  sum_math=sum_math+math[j]
  i=i+5
  j=j+1
#print(math,end="    ")
#print(sum_math)
i=1
j=0
phy=[]
sum_phy=0
while i<len(df) :
  phy.append(df.iloc[i, 8])
  sum_phy=sum_phy+phy[j]
  i=i+5
  j=j+1
#print(phy)
i=2
j=0
chem=[]
sum_chem=0
while i<len(df) :
  chem.append(df.iloc[i, 8])
  sum_chem=sum_chem+chem[j]
  i=i+5
  j=j+1
#print(sum_chem)
i=3
j=0
eng=[]
sum_eng=0
while i<len(df) :
  eng.append(df.iloc[i, 8])
  sum_eng=sum_eng+eng[j]
  i=i+5
  j=j+1
#print(sum_chem)
i=4
j=0
comp=[]
sum_comp=0
while i<len(df) :
  comp.append(df.iloc[i, 8])
  sum_comp=sum_comp+comp[j]
  i=i+5
  j=j+1
#print(sum_comp)
#print(df)
#def grade_edit(x):
  #print(x)
#grade_edit(comp)
  #j=j+1
#for i in range(math):
    #print(math[i])
def rangegrade(mt,j):
   print("1.Enter range for A ")
   print("2.Enter range for B")
   print("3.Enter range for C")
   print("4.Enter range for D")
   print("5.Enter range for E")
   k=int(input())
   if k==1:
       print("HIGH RANGE")
       highmath = int(input())
       print(("LOW RANGE"))
       lowmath = int(input())
       i = 0
       while i < 30:
           if mt[i] >= lowmath and mt[i] < highmath:
               df.at[j, 'Grade'] = 'A'
           j = j + 5
           i = i + 1
   elif k==2:
       print("HIGH RANGE")
       highmath = int(input())
       print(("LOW RANGE"))
       lowmath = int(input())

       i = 0
       while i < 30:
           if mt[i] >= lowmath and mt[i] < highmath:
               df.at[j, 'Grade'] = 'B'
           j = j + 5
           i = i + 1
   elif k==3:
       print("HIGH RANGE")
       highmath = int(input())
       print(("LOW RANGE"))
       lowmath = int(input())
       i = 0
       while i < 30:
           if mt[i] >= lowmath and mt[i] < highmath:
               df.at[j, 'Grade'] = 'C'
           j = j + 5
           i = i + 1
   elif k==4:
       print("HIGH RANGE")
       highmath = int(input())
       print(("LOW RANGE"))
       lowmath = int(input())
       i = 0
       while i < 30:
           if mt[i] >= lowmath and mt[i] < highmath:
               df.at[j, 'Grade'] = 'D'
           j = j + 5
           i = i + 1
   elif k==5:
       print("HIGH RANGE")
       highmath = int(input())
       print(("LOW RANGE"))
       lowmath = int(input())
       i = 0
       while i < 30:
           if mt[i] >= lowmath and mt[i] < highmath:
               df.at[j, 'Grade'] = 'F'
           j = j + 5
           i = i + 1




   #writer.save()
def grademath(summath,mth):
    #print("afbajf")
    devi=0
    mean = summath / 30
    for i in range(0, 30):
        devi = devi + pow(mth[i] - mean, 2)
    devi = pow(devi/30,0.5)
    i = 0
    j = 0
    while i < 30:
        if mean + 2 * devi <= mth[i] < mean + 3 * devi:
            df.at[j, 'Grade'] = 'A'
        elif mth[i] < 35:
            df.at[j, 'Grade'] = 'F'
        elif mean + devi <= mth[i] < mean + 2 * devi:
            df.at[j, 'Grade'] = 'B'
        elif mean - devi <= mth[i] < mean + devi:
            df.at[j, 'Grade'] = 'C'
        elif 35 <= mth[i] < mean - devi:
            df.at[j, 'Grade'] = 'D'
        j=j+5
        i=i+1
grademath(sum_math,math)
def gradephy(sumphy,pth):
    #print("afbajf")
    devi=0
    mean = sumphy / 30
    for i in range(0, 30):
        devi = devi + pow(pth[i] - mean, 2)
    devi = pow(devi/30,0.5)
    i = 0
    j = 1
    while i < 30:
        if mean + 2 * devi <= pth[i] < mean + 3 * devi:
            df.at[j, 'Grade'] = 'A'
        elif pth[i] < 35:
            df.at[j, 'Grade'] = 'F'
        elif mean + devi <= pth[i] < mean + 2 * devi:
            df.at[j, 'Grade'] = 'B'
        elif mean - devi <= pth[i] < mean + devi:
            df.at[j, 'Grade'] = 'C'
        elif 35 <= pth[i] < mean - devi:
            df.at[j, 'Grade'] = 'D'
        j=j+5
        i=i+1
gradephy(sum_phy,phy)
def gradechem(sumchem,cth):
    #print("afbajf")
    devi=0
    mean = sumchem / 30
    for i in range(0, 30):
        devi = devi + pow(cth[i] - mean, 2)
    devi = pow(devi/30,0.5)
    i = 0
    j = 2
    while i < 30:
        if mean + 2 * devi <= cth[i] < mean + 3 * devi:
            df.at[j, 'Grade'] = 'A'
        elif cth[i] < 35:
            df.at[j, 'Grade'] = 'F'
        elif mean + devi <= cth[i] < mean + 2 * devi:
            df.at[j, 'Grade'] = 'B'
        elif mean - devi <= cth[i] < mean + devi:
            df.at[j, 'Grade'] = 'C'
        elif 35 <= cth[i] < mean - devi:
            df.at[j, 'Grade'] = 'D'
        j=j+5
        i=i+1
gradechem(sum_chem,chem)
def gradeeng(sumeng,eth):
    #print("afbajf")
    devi=0
    mean = sumeng / 30
    for i in range(0, 30):
        devi = devi + pow(eth[i] - mean, 2)
    devi = pow(devi/30,0.5)
    i = 0
    j = 3
    while i < 30:
        if mean + 2 * devi <= eth[i] < mean + 3 * devi:
            df.at[j, 'Grade'] = 'A'
        elif eth[i] < 35:
            df.at[j, 'Grade'] = 'F'
        elif mean + devi <= eth[i] < mean + 2 * devi:
            df.at[j, 'Grade'] = 'B'
        elif mean - devi <= eth[i] < mean + devi:
            df.at[j, 'Grade'] = 'C'
        elif 35 <= eth[i] < mean - devi:
            df.at[j, 'Grade'] = 'D'
        j=j+5
        i=i+1
gradeeng(sum_eng,eng)
def gradecomp(sumcomp,coth):
    #print("afbajf")
    devi=0
    mean = sumcomp / 30
    for i in range(0, 30):
        devi = devi + pow(coth[i] - mean, 2)
    devi = pow(devi/30,0.5)
    i = 0
    j = 4
    while i < 30:
        if mean + 2 * devi <= coth[i] < mean + 3 * devi:
            df.at[j, 'Grade'] = 'A'
        elif coth[i] < 35:
            df.at[j, 'Grade'] = 'F'
        elif mean + devi <= coth[i] < mean + 2 * devi:
            df.at[j, 'Grade'] = 'B'
        elif mean - devi <= coth[i] < mean + devi:
            df.at[j, 'Grade'] = 'C'
        elif 35 <= coth[i] < mean - devi:
            df.at[j, 'Grade'] = 'D'
        j=j+5
        i=i+1
gradecomp(sum_comp,comp)
def Edit(t,k):
    print("ENTER NEW MARKS")
    a=int(input())
    df.iloc[t,k]=a
ok=1
while ok:
   print("1. PRINT ALL RECORDS")
   print("2. CHANGE RANGE OF RECORDS")
   print("3. BROWSE STUDENT RECORDS")
   print("4. QUIT PROGRAM")
   q = int(input())
   if q == 1:
      grademath(sum_math, math)
      gradephy(sum_phy, phy)
      gradechem(sum_chem, chem)
      gradeeng(sum_eng, eng)
      gradecomp(sum_comp, comp)
      with pd.option_context('display.max_rows', None, 'display.max_columns', None):  # more options can be specified also
          print(df)
      print("BACK TO MAIN MENU ( Y/N) : ")
      res = input()
      if res =="N" and "n":
          exit()
      if res=="Y" and "y":
         ok=1
   elif q==2:
      print("SELECT SUBJECT")
      print("1. MATHS")
      print("2. PHY")
      print("3. CHEM")
      print("4. ENG")
      print("5. COMP")
      t=int(input())
      if t==1:
        rangegrade(math,0)
      elif t==2:
        rangegrade(phy,1)
      elif t==3:
        rangegrade(chem,2)
      elif t==4:
        rangegrade(eng,3)
      elif t==5:
        rangegrade(comp,4)
      print("BACK TO MAIN MENU ( Y/N) : ")
      res = input()
      if res == "N" and "n":
         exit()
      if res == "Y" and "y":
        ok = 1
   elif q==3:
      print("Enter Student Roll No.")
      l=int(input())
      print("********MARKSHEET OF STUDENT*********")
      print(df.loc[l-1:l+3])
      print("1. EDIT STUDENT RECORD")
      print("2. GO TO MAIN MENU")
      t=int(input())
      if t==1:
         print("SELECT SUBJECT")
         print("1. MATH")
         print("2. PHYSICS")
         print("3. CHEMISTRY")
         print("4. ENGLISH")
         print("5. COMPUTER")
         k=int(input())
         print("1. SESSIONAL1")
         print("2. SESSIONAL1")
         print("3. END SEM")
         print("4. LAB")
         n = int(input())
         Edit(k-1,n+3)
         print("BACK TO MAIN MENU ( Y/N) : ")
         res = input()
         if res == "N" and "n":
            exit()
         if res == "Y" and "n":
           ok = 1
      elif t==2:
         ok = 1
   elif q==4:
       exit()
    #Exit
writer = pd.ExcelWriter('std4.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1')
writer.save()
# with pd.option_context('display.max_rows', None, 'display.max_columns', None):  # more options can be specified also
#    print(df)