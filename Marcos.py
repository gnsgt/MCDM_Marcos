import pandas as pd
import numpy as np

    #Excel dosyasını oku ve DataFrame'e dönustur
df_excel = pd.read_excel("C:/Users/gokha/OneDrive/Masaüstü/Tez/11.1-Marcos.xlsx")
maliyet = df_excel.iloc[1:2].copy()
agirlik = df_excel.iloc[:1].copy()
df = df_excel.iloc[2:]  
satir_sayisi = df.shape[0] 
sutun_sayisi = df.shape[1]


aal = []
al = []
for i in range(sutun_sayisi):
        if maliyet.iloc[0,i] == "fayda":
            aal.append(min(df.iloc[0:,i]))
            al.append(max(df.iloc[0:,i]))
        else:
            aal.append(max(df.iloc[0:,i]))
            al.append(min(df.iloc[0:,i]))
           

normalize_df = df.copy()
for j in range(0,sutun_sayisi):
    for i in range(0,satir_sayisi):
        if maliyet.iloc[0,j] == "fayda": 
            normalize_df.iloc[i,j] = df.iloc[i,j] / al[j]
        else:
            normalize_df.iloc[i,j] = al[j] /  df.iloc[i,j]
            

            
for i in range(0,sutun_sayisi):
    al[i] = max(normalize_df.iloc[:,i])
    aal[i] = min(normalize_df.iloc[:,i])

wndf = normalize_df.copy()
for i in range(0,sutun_sayisi):
    wndf.iloc[:,i] = normalize_df.iloc[:,i] * agirlik.iloc[0,i]
    aal[i] = aal[i] * agirlik.iloc[0,i]
    al[i] = al[i] * agirlik.iloc[0,i] 


saai = sum(aal)
sai = sum(al)

si = []
for i in range(0,satir_sayisi):
    si.append(sum(wndf.iloc[i,:]))

minuski = []
pluski = []


for i in range(len(si)):
    minuski.append(si[i] / saai)
    pluski.append(si[i] / sai)


fminuski = []
fpluski = []
for i in range(len(si)):
    fminuski.append(pluski[i] / (pluski[i]+minuski[i]))
    fpluski.append(minuski[i] / (minuski[i]+pluski[i]))
    
fki = []
for i in range(len(si)):
     fki.append((pluski[i] + minuski[i]) / (1+((1-fpluski[i])/fpluski[i])+((1-fminuski[i])/fminuski[i])))
     
print("\n")
for i in range(satir_sayisi):
    max_deger = max(fki)
    max_index = fki.index(max_deger)
    print(f"Marcos Ki: {max_deger}, İndex: {max_index + 1}")
    fki[max_index] = -99999

