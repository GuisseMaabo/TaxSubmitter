from lxml import etree 
import pandas as pd 


#importing excel file
file = "decla.xlsx"
df = pd.read_excel(file, sheet_name='T1 Bilan')
# replace nan by 0
df_nan = df.fillna(0)
#print(df_nan)
# selecting specific  rows and columns 
newdf = df_nan.iloc[8:,2:]
#print(newdf) 

# testing the loop 
"""
for rowo in newdf.iterrows():
    print(str(rowo[1]['Unnamed: 2']))
    print(str(rowo[1]['Unnamed: 3']))
    print(str(rowo[1]['Unnamed: 4']))
    print(str(rowo[1]['Unnamed: 5']))


"""


# Immobilisation en non valeur A 
df = pd.DataFrame([[497,498,499,500],
                [238,241,244, 247],
                [239,242,245, 248],
                [240,243,246,249]],
                columns=['Brut exercice', 'Amortissements et provisions : exercice', 'Net exercice', 'Net exercice précédent'],
                index = ['IMMOBILISATION EN NON-VALEURS(A)', '* Frais préliminaires','* Charges à répartir sur plusieurs exercices',  '* Primes de remboursement obligations'])
#print(df)

# Immobilisation en non valeurs B 
df1 = pd.DataFrame([[502,503,504,505],
                [274,278,282,286],
                [275,279,283, 287],
                [276,280,284,288],
                [277,281,285,289]],
                columns=['Brut exercice','Amortissements et provisions : exercice', 'Net exercice', 'Net exercice précédent'],
                index = ['IMMOBILISATION EN NON-VALEURS(B)','Immobilisation en recherche et développement','Brevets','Fond commercial','Autres immobilisations incorporelles'])
#print(df1)
dfi = pd.concat([df,df1], axis=0)
#print(dfi)
# Immobilisation Corporelles C 
df2 = pd.DataFrame([[507,508,509,510],
                [207,214,221,228],
                [208,215,222, 229],
                [209,216,223,230],
                [210,217,224,231],
                [211,218,225,232],
                [212,219,226,233],
                [213,220,227,234],],
                columns=['Brut exercice','Amortissements et provisions : exercice', 'Net exercice', 'Net exercice précédent'],
                index = ['IMMOBILISATION EN NON-VALEURS(C)','Terrains','Constructions','Installations techninques, materiel et outillage','Materiel de transport', 'Mobilier, Materiel de bureau et aménagement divers', 'Autres immobilisation corporelles', 'Immobilisations corporelles en cours'])
dfi1 = pd.concat([dfi,df2], axis=0)
#print(dfi1)

# Immobilisations financieres (D)
dfi2 = pd.DataFrame([[512,513,514,515],
                [254,258,262,266],
                [255,259,263, 267],
                [256,260,264,268],
                [257,261,265,269]],
                columns=['Brut exercice','Amortissements et provisions : exercice', 'Net exercice', 'Net exercice précédent'],
                index = ['Immobilisation financière (D)','Prêts immobilisés','Autres créances financières','Titres de participation','Autres immobilisés'])
dfi3 = pd.concat([dfi1,dfi2], axis=0)
#print(dfi3)


# Ecarts de conversion- Actif(E)
dfi4 = pd.DataFrame([[517,518,519,520],
                [192,194,196, 198],
                [193,195,197,199],
                [291,292,293,294],],
                columns=['Brut exercice', 'Amortissements et provisions : exercice', 'Net exercice', 'Net exercice précédent'],
                index = ['Ecarts de conversion actif (E)', 'Diminution des créances immobilisées','Augmentation des dettes de financement', 'Total I (A + B + C + D + E )'])
dfi5 = pd.concat([dfi3,dfi4], axis=0)
#print(dfi5)
# Stock F
dfi6 = pd.DataFrame([[522,523,524,525],
                [160,165,170, 175],
                [161,166,171,176],
                [162,167,172,177],
                [163,168,173,178],
                [164,169,174,179],],
                columns=['Brut exercice', 'Amortissements et provisions : exercice', 'Net exercice', 'Net exercice précédent'],
                index = ['Stock (F)', 'Marchandises','Matiéres et fournitures consommable', 'Produits en cours', 'Produits intermediaires et Produits résiduels', 'Produits finis'])
dfi7 = pd.concat([dfi5,dfi6], axis=0)
#print(dfi7)

# Creances de l'actif circulant G 
dfi8 = pd.DataFrame([[527,528,529,530],
                [122,129,136,143],
                [123,130,137,144],
                [124,131,138,145],
                [125,132,139,146],
                [126,133,140,147],
                [127,134,141,148],
                [128,135,142,149],],
                columns=['Brut exercice','Amortissements et provisions : exercice', 'Net exercice', 'Net exercice précédent'],
                index = ['Créance de l\'actif circulant(G)','Fournis, debiteurs, avances et acomptes','Clients et comptes ratachés','Personnel','Etat', 'Comptes d\'associés', 'Autres debiteurs' , 'comptes et regularisation-actif'])
dfi9 = pd.concat([dfi7,dfi8], axis=0)
print(dfi9)


# Creating my Xml file
root = etree.Element("Liasse")

for rowo in newdf.iterrows():
    for row in dfi9.iterrows():
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
#print(etree.tostring(root, pretty_print=True))



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



