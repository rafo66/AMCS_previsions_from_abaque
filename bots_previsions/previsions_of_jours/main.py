'''
1. Prendre les listes des OF/Machines de SAP
2. Trouver chaques ref
3. Trouver les prods correspondents
4. Extraire le nombre de postes aujourd'hui
5. Prédire le nombre d'OF réalisé en fonction du nombre  de postes 
'''
import os 
import pandas as pd
import math

class previsions():

    def __init__(self, ligne="R6", duree=7.5):
        self.req_ligne = ligne
        self.req_duree = duree

        pass

    def get_of_par_machines(self):
        xlsxFile = os.path.join(os.getcwd(), "liste_of_ligne", self.req_ligne + ".xlsx")
        df = pd.read_excel(xlsxFile)

        '''
        ['OF Principal', 'N° OF', 'Article1', 'Statut du lot', 'Client1',
        'FW Plan de conditionnement', 'FW Format (..x..x..)', 'Nb de bandes',
        'Stock PM', 'Désignation article1', 'Ordre1', 'Prématériel',
        'VM Sous-catégorie produit', 'Prématériel.1', 'Client1.1']
        '''
        
        #print(df.head())
        return df
    

    def level_0(self, article, abaqueDF):
        #print("LEVEL 0: Exact Match Article", int(article))
        candidats = []

        for index, row in abaqueDF.iterrows():
            try:
                if str(int(article)) in str(row["Articles"]):
                    candidats.append(row["Prod T/h/OF"])
            except:
                pass

        if len(candidats) > 0:
            return sum(candidats)/len(candidats)
        return -1
    
    def find_prod(self, _of_line):
        xlsxFile = os.path.join(os.getcwd(), "Abaque", "Abaque.xlsm")
        abaqueDF = pd.read_excel(xlsxFile)
        #print(abaqueDF.columns)

        '''
        'OF', 'Prod T/h/OF', 'Bandes (/OF)', 'Scindage (/OF)', 'Date 1',
        'Temps 1 (h)', 'Année 1', 'Tonnes', 'heures', 'Épaisseur Nominal',
        'Largeur', 'Longueur', 'Proto', 'Poste', 'Nb pieces', 'Date 2',
        'Temps 2 (h)', 'Année 2', 'Clients', 'Articles']
        
        '''

        '''
        ' LEVEL 0: Exact Match Article            
        ' LEVEL 1: Exact Match (Name + Proto + 3 Dims)
        ' LEVEL 2: Partial Match (Name + Proto + Dim1)
        ' LEVEL 3: Partial Match (Name + Proto)
        ' LEVEL 4: Moyenne ligne
        '''
        
        return (self.level_0(_of_line["Article1"], abaqueDF))
    
    def get_postes_par_ligne(self):
        return 2
    
    def get_requiered_time(self, article):
        prod = self.find_prod(article)
        return article["poids"] / prod
    
    def globalProc(self):
        req = self.get_of_par_machines()
        
        for index, row in req.iterrows():
            print(index, self.find_prod(row))
    

p=previsions()
p.globalProc()