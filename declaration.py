from lxml import etree 
import pandas as pd 


#importing excel file
file = "declaration.xlsx"
df = pd.read_excel(file, sheet_name='T1 Bilan')
# replace nan by 0
df_nan = df.fillna(0)
print(df_nan)
# selecting specific  rows and columns 
newdf = df_nan.iloc[7:,2:]
print(newdf) 

# testing the loop 

for rowo in newdf.iterrows():
    print(str(rowo[1]['Unnamed: 2']))
    print(str(rowo[1]['Unnamed: 3']))
    print(str(rowo[1]['Unnamed: 4']))
    print(str(rowo[1]['Unnamed: 5']))




df = pd.DataFrame([[10064, 966, 967, 968],
                [10065, 980, 981, 982],
                [10066, 994, 995, 996],
                [10067, 1008 , 1009, 1010]],
                columns=['Brut exercice', 'Amortissements et provisions : exercice', 'Net exercice', 'Net exercice précédent'],
                index = ['IMMOBILISATION EN NON-VALEURS', '* Frais préliminaires','* Charges à répartir sur plusieurs exercices',  '* Primes de remboursement obligations'])

print(df)


# Creating my Xml file
root = etree.Element("Liasse")

for rowo in newdf.iterrows():
    for row in df.iterrows():
        valeurTableau = etree.SubElement(root,'ValeurTableau')
        tableau = etree.SubElement(valeurTableau,'Tableau')
        valeurCellule = etree.SubElement(tableau,'ValeurCellule')
        ####################### Brut exercice ###############################
        cellule = etree.SubElement(valeurCellule,'Cellule')
        valeur = etree.SubElement(cellule,'valeur')
        valeur.text = str(rowo[1]['Unnamed: 2'])
        codeEdi = etree.SubElement(cellule,'codeEdi')
        codeEdi.text = str(row[1]['Brut exercice'])
        #################### Amortissements et provisions : exercice ###########################
        cellule = etree.SubElement(valeurCellule,'Cellule')
        valeur = etree.SubElement(cellule,'valeur')
        valeur.text =str(rowo[1]['Unnamed: 3'])
        codeEdi = etree.SubElement(cellule,'codeEdi')
        codeEdi.text = str(row[1]['Amortissements et provisions : exercice'])
        ################## Net exercice #############################
        cellule = etree.SubElement(valeurCellule,'Cellule')
        valeur = etree.SubElement(cellule,'valeur')
        valeur.text = str(rowo[1]['Unnamed: 4'])
        codeEdi = etree.SubElement(cellule,'codeEdi')
        codeEdi.text = str(row[1]['Net exercice'])
        ##################### Net exercice précédent ##########################
        cellule = etree.SubElement(valeurCellule,'Cellule')
        valeur = etree.SubElement(cellule,'valeur')
        valeur.text = str(rowo[1]['Unnamed: 5'])
        codeEdi = etree.SubElement(cellule,'codeEdi')
        codeEdi.text = str(row[1]['Net exercice précédent'])
print(etree.tostring(root, pretty_print=True))



with open('doc.xml', 'w') as f:
    f.write(etree.tostring(root, pretty_print=True ).decode('utf-8'))
    f.close() 

   


"""
parser = ET.XMLParser(remove_blank_text=True)
tree = ET.parse(xml, parser)
tree.write(xml, encoding='utf-8', pretty_print=True, xml_declaration=True)
"""



# Create an xml file 
"""
f = open('doc.xml', 'w')
f.write(ET.tostring(root, pretty_print=True))
f.close()
"""



