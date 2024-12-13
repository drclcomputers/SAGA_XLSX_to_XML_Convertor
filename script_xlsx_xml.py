import pandas as pd
import xml.etree.ElementTree as ET

"""
def isnumeric(s): 
    try: 
        int(s) 
        return True 
    except: 
        return False 
        
def vf_zi(zi):
    while not isnumeric(zi):
        zi=int(input("Introdu o zi valida: "))
    while not (zi<1 or zi>31):
        zi=int(input("Introdu o zi valida: "))
    return zi
        
def vf_luna(dat):
    while not isnumeric(dat):
        dat=int(input("Introdu o luna valida: "))
    while not (luna<1 or luna>12):
        dat=int(input("Introdu o luna valida: "))
    return luna
        
def vf_an(an):
    while not isnumeric(an):
        an=int(input("Introdu un an valida: "))
    while not (an<2000 or an>2100):
        an=int(input("Introdu un an valida: "))
    return an
"""


locatie=input("Introdu adresa fisierului xlsx: ")

try:
    df = pd.read_excel(locatie)
except:
    print("Fisier inexistent!")
    exit()
    
    

zi=input("Introdu ziua: ")
#zi=vf_zi(zi)
dat=input("Introdu luna: ")
#dat=vf_luna(dat)
an=input("Introdu anul: ")
#an=vf_an(an)



root = ET.Element("Articole_Contabile")

i=1
aux=1

for _, row in df.iterrows():

    dat_aux=str(row["DATA"])
    
    if dat_aux[5:7]==dat:
        entry = ET.SubElement(root, "Linie")
        
        data=str(row["DATA"])
        
        ET.SubElement(entry, "Data").text = data[:-9]
        
        ET.SubElement(entry, "Numar").text = str(i)
        
        ET.SubElement(entry, "Suma").text = str(row["SUMA"])
        
        ET.SubElement(entry, "Explicatie").text = str(row["EXPLICAȚII"])
        
        if aux==2:
            i=i+1
            aux=0
            
        aux=aux+1

tree = ET.ElementTree(root)
tree.write(str("NC_01-01-2017_31-12-2017_"+zi+"-"+dat+"-"+an+".xml"), encoding="utf-8", xml_declaration=True)