# -*- coding: utf-8 -*-
"""
Created on Sat Apr 28 15:12:05 2018

@author: Vinod Kumar Tudu
"""

import pandas as pd
import numpy as np
import os
inv=pd.read_excel('invoice.xlsx')
pdf=pd.read_excel('PPD.xlsx')
cdf=pd.read_excel('COD.xlsx')
print(inv.head(1))

PPD=-1
COD=-1
inv1=0
inv2=0
finalpd=pd.DataFrame(columns=['AWB_NUMBER'	,'ORDER_NUMBER'	,'PRODUCT'	,'CONSIGNEE'	,'CONSIGNEE_ADDRESS1'	,'DESTINATION_CITY'	,'PINCODE'	,'STATE'	,'MOBILE'	,'ITEM_DESCRIPTION'	,'PIECES',	'COLLECTABLE_VALUE',	'DECLARED_VALUE'	,'INVOICE_NUMBER',	'INVOICE_DATE	SELLER_GSTIN'	,'GST_TAX_NAME'	,'GST_TAX_BASE'	,'GST_TAX_TOTAL'])

#sub invoice
def finddetail(inv):
    global COD,PPD,pdf,cdf
    
    a=search(inv,'Mode/Terms of Payment')
    if 'COD' in a:
        cod=1
        COD=COD+1
        AWB=cdf.loc[COD][0]
        cdf.loc[COD][1]='used'
    else :
        cod=0
        PPD=PPD+1
        AWB=pdf.loc[PPD][0]
        pdf.loc[PPD][1]='used'
    orderno=search(inv,"Buyer's Order Number")
    if cod==1:
        product='COD'
    else:
        product='PPD'
        
    i,j,a=deepsearch(inv,'state code')
    statecode=inv.loc[i][j+1]
    state=inv.loc[i-1][j+1]
    mobile=inv.loc[i-2][j].split(':')[1]
    address=inv.loc[i-3][j]
    name=inv.loc[i-4][j]
    city=address.split(',')[-1]
    data=(AWB,orderno,product,name,address,city,'<pincode>',state,mobile)
    return data
    
    
def readitems(inv,data1):
    data=list(data1)
    global finalpd
    final=[]
    i,j,a=deepsearch(inv,'Description of Goods')
    i=i+1
    x1,y1,z=deepsearch(inv,'total')
    total=inv.loc[x1][y1+8]
    while (str(inv.loc[i][j]).lower()!='nan'):
        data=list(data1)
        data.append(inv.loc[i][j])
        data.append(inv.loc[i][j+5])
        #total=inv.loc[i][j+8]
        if data[2]!='COD':
            data.append(0)
        else :
            data.append(total)
        data.append(total)
        data.append(search(inv,'Invoice Number'))
        data.append(search(inv,'Dated'))
        x1,y1,z=deepsearch(inv,'GSTIN')
        gst=z.split(' ')[1]
        data.append(gst)
        x1,y1,z=deepsearch(inv,'total')
        gsttype=inv.loc[x1-1][y1+7]
        data.append('HR '+gsttype)
        total=inv.loc[i][j+8]
        data.append(total)
        data.append(total*.05)
        i=i+1
        print (type(data))
        finalpd=finalpd.append(pd.Series(data[1:],index=['AWB_NUMBER'	,'ORDER_NUMBER'	,'PRODUCT'	,'CONSIGNEE'	,'CONSIGNEE_ADDRESS1'	,'DESTINATION_CITY'	,'PINCODE'	,'STATE'	,'MOBILE'	,'ITEM_DESCRIPTION'	,'PIECES',	'COLLECTABLE_VALUE',	'DECLARED_VALUE'	,'INVOICE_NUMBER',	'INVOICE_DATE	SELLER_GSTIN'	,'GST_TAX_NAME'	,'GST_TAX_BASE'	,'GST_TAX_TOTAL']), ignore_index=True)
        #finalpd=finalpd.append(data)
    return final
        
        
        
def deepsearch(inv,word):
    for i in range(len(inv)):
        for j in range(len(inv.iloc[0])):
            a=str(inv.loc[i][j])
            if word.lower() in a.lower():
                return i,j,a
def search(inv,query):
    for i in range(len(inv)):
        for j in range(len(inv.iloc[0])):
            a=str(inv.loc[i][j])
            if a==query:
                return inv.loc[i+1][j]

def read_head(inv):
    comp=[]
    mark=[]
    global inv1,inv2
    a=list(inv.iloc[0])
    for i in range(len(a)):
        j=str(a[i])
        #print (j,type(j))
        if j.lower()!='nan' :
            mark.append(i)
            comp.append(a[i])
    #print (comp)
    
    splitlen=mark[1]-mark[0]
    inv1=inv.iloc[:, :splitlen-1]
    inv2=inv.iloc[:,splitlen:]
    print(splitlen)
    data1=finddetail(inv1)
    data2=finddetail(inv2)
    #print (data1)
    #print(data2)
    final1=readitems(inv1,data1)
    final2=readitems(inv2,data2)
comp=read_head(inv)
finalpd.to_excel('output.xlsx',index=False)
pdf.to_excel('PPD.xlsx',index=False)
cdf.to_excel('COD.xlsx',index=False)

for root,dirs,files in os.walk('.'):
        for file in files:
            #print (file)
            if file.endswith('.xlsx'):
                fullpathname=os.path.join(root,file)
                #print (fullpathname)
                filename=fullpathname.split("\\")
                #print (filename)
                if len(filename)<3:
                    print (len(filename),filename[-1].lower())
                if len(filename)<3 and 'invoice' in filename[-1].lower():
                    print ('reading - ',filename[-1] )
                    inv=pd.read_excel(filename[-1])
                    comp=read_head(inv)
                    print (finalpd.shape)
                    
finalpd.to_excel('output.xlsx',index=False)
pdf.to_excel('PPD.xlsx',index=False)
cdf.to_excel('COD.xlsx',index=False)