'''
1. Prendre les listes des OF/Machines de SAP
2. Trouver chaques ref
3. Trouver les prods correspondents
4. Extraire le nombre de postes aujourd'hui
5. Prédire le nombre d'OF réalisé en fonction du nombre  de postes 



Todo : Trouver le poids, depuis SAP liste OF / machine

'''
import os
import time 
import pandas as pd
import math

oldLines =     ["D10", "D10R11", "D14R11", "D20", "D7R10", "FIMI", "FIMIR3B", "L1", "LAS1.", "P3", "R1", "R10", "R11", "R2", "R3B", "R6", "R7", "P3R1"]
newLines =     ["L1", "L1",      "L1",      "L1", "L1",     "L1",  "L1",      "L1", "LASS1,", "P3", "R1", "R1", "R6", "R6", "R6", "R6", "R1", "P3"]
avgLineProto = [4.15, 4.15,      4.15,      4.15, 4.15,     4.15, 4.15,        4.15, 0.39,     2.25, 14,   14,   15.8, 15.8, 15.8, 15.8, 14, 2.25]
avgLineSerie = [6.96, 6.69,      6.69,      6.96, 6.96,     6.96, 6.96,        6.96,  0.7,      2.5, 23.2, 23.2, 39,   39,   39,   39,   23.2,  2.5]

class MatchingProductivities:
    '''
        Takes 2 dataframes as input :
        - Prevision_ligne_Df : dataframe of the exception report, with only Woippy plant data
        - abaqueDf : dataframe of the abaque, with all the data

        Returns the exceptionReportDf with a new column "Productivity" filled with the matched productivity for each row, and a new column "Match Level" filled with the level of the match (0 to 4, or -1 if no match)

        Exemple : 
        
        mP = MatchingProductivities(self.getPrevision_ligne_Df(), self.getAbaqueDf(), self.cacheFile)
        self.newExceptionReportDf = mP.exceptionReportDf
        
        if cacheFile is empty or None, no caching will be done. Else articleCachedProductivities.txt is used

        Dependencies : 
            - OS access to read and write cache file
            - pandas for dataframe manipulation


        

        ' LEVEL 0: Exact Match Article            
        ' LEVEL 1: Exact Match (Name + Proto + 3 Dims)
        ' LEVEL 2: Partial Match (Name + Proto + Dim1)
        ' LEVEL 3: Partial Match (Name + Proto)
        ' LEVEL 4: Moyenne ligne

    '''
            
    def __init__(self, Prevision_ligne_Df, abaqueDf, ligne, cacheFile="articleCachedProductivities"):
        self.exceptionReportDf = Prevision_ligne_Df
        self.abaqueDf = abaqueDf
        self.ligne = ligne
        self.cacheFile = cacheFile + "_" + ligne + ".txt"

        self.articleCachedProductivities = self.loadArticleCachedProductivities()

        self.matchProductivities()

    
    def loadArticleCachedProductivities(self):
        if self.cacheFile == None or self.cacheFile == "":
            return {}
        
        # Load cached productivities from a file if it exists, otherwise return an empty dictionary
        if os.path.exists(self.cacheFile):
            with open(self.cacheFile, "r") as f:
                lines = f.readlines()
                articleCachedProductivities = {}
                for line in lines:
                    article, prod, details, level = line.strip().split(":")
                    articleCachedProductivities[int(article)] = [float(prod), details, int(level)]
                print("Loaded cached productivities from file: ", articleCachedProductivities)
                return articleCachedProductivities
        else:
            return {}
        
    def saveCachedProductivities(self):
        if self.cacheFile == None or self.cacheFile == "":
            return
        
        
        # Save cached productivities to a file
        with open(self.cacheFile, "w") as f:
            for article, prod in self.articleCachedProductivities.items():
                f.write(f"{int(article)}:{round(float(prod[0]), 2)}:{prod[1]}:{prod[2]}\n")


    def matchProductivities(self):
        '''
        Exception Report Columns: Sales Document | Material | Plant | Last Updated 
        | Sold-to party | Name of sold-to party | Price Partner | Name of Price partner 
        | Ship-to Party | Name of ship-to party | Ship-To Location | External office code 
        | Name of External Office | Back Office code | Name of Back Office | Sales Type 
        | Sales unit | Numerator | Denominator | Number of Pieces | Call-Offs | Sold to Article Number 
        | Aramis Grade | Thickness | Width | Length | Base Unit of Measure | Customer Order Reference 
        | Internal Order Number | Material for BOM | Critical Part | Mill | Leverage contract list 
        | Leverage Article list | BOM Alert | Routing | Backlog | Forecast W | Forecast W1 | Forecast W2 
        | Forecast W3 | Forecast W4 | Forecast W5 | Forecast W6 | Forecast W7 | Forecast W8 | Forecast W-1 
        | Forecast W-2 | Aramis Open Quantity_W | Aramis Open Quantity_W1 | Aramis Open Quantity_W2 
        | Aramis Open Quantity_W3 | Aramis Open Quantity_W4 | Aramis Open Quantity_W5 | Aramis Open Quantity_W6 
        | Aramis Open Quantity_W7 | Aramis Open Quantity_W8 | Aramis Open Quantity_D0 | Aramis Open Quantity_D1 
        | Aramis Open Quantity_D2 | Aramis Open Quantity_D3 | Aramis Open Quantity_D4 | Aramis Open Quantity_D5 
        | Aramis Open Quantity_D6 | Aramis Open Quantity_D7 | Aramis Open Quantity_D8 | Aramis Open Quantity_D9 
        | Aramis Open Quantity_D10 | Aramis Open Quantity_D11 | Aramis Open Quantity_D12 | Aramis Open Quantity_D13 
        | Aramis Open Quantity_D14 | Aramis Open Quantity_D15 | Call-Offs_W | Call-Offs_W1 | Call-Offs_W2 | Call-Offs_W3 
        | Call-Offs_W4 | STOCK_FG_FREE | Blocked Quality | Qual Inspection | Consignment Stock | On a transport request 
        | Unrestricted RM | Stock in tfr RM | Quality Inspection RM | Restricted-Use Stock RM | Blocked RM 
        | Unrestr. Consignment RM | Cnsgt in Inspection RM | Restr. Consignment RM | Blocked Consignment RM 
        | RTS at the Mill | RTS at the Mill UOM | In Transit from the Mill | In Transit from the Mill UOM 
        | Backlog Detail W-0 | Backlog Detail W-1 | Backlog Detail W-2 | Backlog Detail W-3 | Backlog Detail W-4 
        | Backlog Detail W-5 | Backlog Detail W-6 | Backlog Detail W-7 | Backlog Detail W-8 | Backlog Details W-9 and more 
        | Backlog Detail D-1 | Backlog Detail D-2 | Backlog Detail D-3 | Backlog Detail D-4 | Backlog Detail D-5 
        | Backlog Detail D-6 | Backlog Detail D-7 | Production : Unprocessed quantity past | Production : Unprocessed quantity W0 
        | Production : Unprocessed quantity W1 | Production : Unprocessed quantity W2 | Production : Unprocessed quantity W3 
        | Production : unplanned quantity | Production : Finish planned in the Past | Finish Planned W-0 | Finish Planned W-1 
        | Finish Planned W-2 | Finish Planned W-3 | Finish Planned W-4 | Unit | Date of the last Call-off Cust Version 
        | Date of the last forecast Cust version | Last Updated week | To be delivered W-0 | To be delivered W-1 
        | To be delivered W-2 | To be delivered W-3 | To be delivered W-4 | To be delivered W-5 | To be delivered W-6 
        | To be delivered W-7 | To be delivered W-8 | Quantity in transit to customer | Total Stock | To be delivered D-0 
        | To be delievred D-1 | To be delievred D-2 | To be delievred D-3 | To be delievred D-4 | To be delievred D-5 
        | To be delievred D-6 | To be delievred D-7 | To be delievred D-8 | RM | Plant_RM | Nb Sales / RM | FG (Total) 
        | WO Plan | WO Non Plan (SO) | WO Non Plan (RM) | RM (Plant + Consi) | RM Available | RM W | RM W+1 | RM W+2 
        | RTB Ship Backlog | RTB Ship W | RTB Ship W+1 | RTB Ship W+2 | RTB Ship W+3 | RTB Ship W+4 | RTB Ship W+5 
        | RTB Ship W+6 | RTB Ship W+7 | RTB Ship W+8 | Lack FG Backlog (SO) | Lack FG W (SO) | Lack FG W+1 (SO) 
        | Lack FG W+2 (SO) | Lack FG+Plan Backlog (SO) | Lack FG+Plan W (SO) | Lack FG+Plan W+1 (SO) | Lack FG+Plan W+2 (SO) 
        | Lack FG+Plan Backlog (RM) | Lack FG+Plan W (RM) | Lack FG+Plan W+1 (RM) | Lack FG+Plan W+2 (RM) 
        | Plan Order Backlog (SO) | Plan Order W (SO) | Plan Order W+1 (SO) | Plan Order W+2 (SO) | Nb Plan Order Backlog (SO) 
        | Nb Plan Order W (SO) | Nb Plan Order W+1 (SO) | Nb Plan Order W+2 (SO) | Plan Order Backlog (RM) | Plan Order W (RM) 
        | Plan Order W+1 (RM) | Plan Order W+2 (RM) | Lack RM Plant Backlog (SO) | Lack RM Plant W (SO) | Lack RM Plant W+1 (SO) 
        | Lack RM Plant W+2 (SO) | Lack RM Plant Backlog (SO).1 | Lack RM Plant W Cumul (SO) | Lack RM Plant W+1 Cumul (SO) 
        | Lack RM Plant W+2 Cumul (SO) | Lack RM Plant Backlog (RM) | Lack RM Plant W (RM) | Lack RM Plant W+1 (RM) 
        | Lack RM Plant W+2 (RM) | Lack RM Plant Backlog (RM).1 | Lack RM Plant W Cumul (RM) | Lack RM Plant W+1 Cumul (RM) 
        | Lack RM Plant W+2 Cumul (RM) | Lack RM AM Backlog | Lack RM AM W | Lack RM AM W+1 | Lack RM AM W+2 | FG Backlog 
        | FG W | FG W+1 | FG W+2 | RM Plant Backlog | RM Plant W | RM Plant W+1 | RM Plant W+2 | RM AM Backlog | RM AM W 
        | RM AM W+1 | RM AM W+2 | Alerte Backlog | Alerte W | Alerte W+1 | Alerte W+2 | GOR Backlog | GOR W | GOR W+1 | GOR W+2 
        | Lack RM Type | Lack RM Cover | GOR TYPE Backlog | GOR TYPE W | GOR TYPE W+1 | GOR TYPE W+2 | Horizon Backlog 
        | Horizon W | Horizon W+1 | Horizon W+2 | No Needs Horizon W+2 | Action to do | Top 10 | Stay to deliver following 2 weeks 
        | total ship | total cut back & week | total call Back & week | FG Available for week ongoing | campagne optimizer  
        | Weekly Coverage FG | Weekly Coverage RM | RM.1 | Weekly Consumption | RTB Ship Cumul Backlog | RTB Ship Cumul W 
        | RTB Ship Cumul W+1 | RTB Ship Cumul W+2 | RTB Ship Cumul W+3 | RTB Ship Cumul W+4 | RTB Ship Cumul W+5 
        | RTB Ship Cumul W+6 | RTB Ship Cumul W+7 | RTB Ship Cumul W+8 | FG | FG Backlog.1 | FG W.1 | FG W+1.1 | FG W+2.1 
        | FG W+3 | FG W+4 | FG W+5 | FG W+6 | FG W+7 | FG W+8 | Plan | Plan Backlog | Plan W | Plan W+1 | Plan W+2 | Plan W+3 
        | Plan W+4 | Plan W+5 | Plan W+6 | Plan W+7 | Plan W+8 | WO | WO Backlog | WO W | WO W+1 | WO W+2 | WO W+3 | WO W+4 
        | WO W+5 | WO W+6 | WO W+7 | WO W+8 | Plan Order Backlog | Plan Order W | Plan Order W+1 | Plan Order W+2 | Plan Order W+3 
        | Plan Order W+4 | Plan Order W+5 | Plan Order W+6 | Plan Order W+7 | Plan Order W+8 | NB PlOr Backlog | NB PlOr W 
        | NB PlOr W+1 | NB PlOr W+2 | NB PlOr W+3 | NB PlOr W+4 | NB PlOr W+5 | NB PlOr W+6 | NB PlOr W+7 | NB PlOr W+8 | RM.2 
        | RM Backlog | RM W.1 | RM W+1.1 | RM W+2.1 | RM W+3 | RM W+4 | RM W+5 | RM W+6 | RM W+7 | RM W+8 | Bes Transit Backlog 
        | Bes Transit W | Bes Transit W+1 | Bes Transit W+2 | Bes Transit W+3 | Bes Transit W+4 | Bes Transit W+5 | Bes Transit W+6 
        | Bes Transit W+7 | Bes Transit W+8 | NB Transit Backlog | NB Transit W | NB Transit W+1 | NB Transit W+2 | NB Transit W+3 
        | NB Transit W+4 | NB Transit W+5 | NB Transit W+6 | NB Transit W+7 | NB Transit W+8 | Transit | Transit Backlog | Transit W 
        | Transit W+1 | Transit W+2 | Transit W+3 | Transit W+4 | Transit W+5 | Transit W+6 | Transit W+7 | Transit W+8 | Bes RTS Backlog 
        | Bes RTS W | Bes RTS W+1 | Bes RTS W+2 | Bes RTS W+3 | Bes RTS W+4 | Bes RTS W+5 | Bes RTS W+6 | Bes RTS W+7 | Bes RTS W+8 
        | NB RTS Backlog | NB RTS W | NB RTS W+1 | NB RTS W+2 | NB RTS W+3 | NB RTS W+4 | NB RTS W+5 | NB RTS W+6 | NB RTS W+7 | NB RTS W+8 
        | RTS | RTS Backlog | RTS W | RTS W+1 | RTS W+2 | RTS W+3 | RTS W+4 | RTS W+5 | RTS W+6 | RTS W+7 | RTS W+8 | NO RM Backlog | NO RM W 
        | NO RM W+1 | NO RM W+2 | NO RM W+3 | NO RM W+4 | NO RM W+5 | NO RM W+6 | NO RM W+7 | NO RM W+8 | Coverage Backlog | Coverage W 
        | Coverage W+1 | Coverage W+2 | Coverage W+3 | Coverage W+4 | Coverage W+5 | Coverage W+6 | Coverage W+7 | Coverage W+8 
        | Perf délais Cumul forecast W-2 + W-1 | Perf délais Cumul forecast W-2 to W+2 | Perf délais  Cumul forecast W-2+ W+8 
        | Perf délais Backlog sup. 0,5 T | Analyse Backlog W-1 | Analyse Forecast w-1 | Transit on FG | Week No Covered by FG 
        | Flag Coverage FG | Week No Covered by Stock SSC | Flag Coverage Stock SSC | RTS Horizon W+8 | Concat PLANT RM | yy | 

        '''
        '''
        Abaque Columns: OF | Prod T/h/OF | Bandes (/OF) | Scindage (/OF) | Date 1 
        | Temps 1 (h) | Année 1 | Tonnes | heures | Épaisseur Nominal | Largeur 
        | Longueur | Proto | Poste | Nb pieces | Date 2 | Temps 2 (h) | Année 2 
        | Clients | Articles |
        '''

        print(self.exceptionReportDf.columns)
        ''' Prevision columns 
        
            OF Principal', 'N° OF', 'Article1', 'Statut du lot', 'Client1',
            'FW Plan de conditionnement', 'FW Format (..x..x..)', 'Nb de bandes',
            'Stock PM', 'Désignation article1', 'Ordre1', 'Prématériel',
            'VM Sous-catégorie produit', 'Prématériel.1', 'Client1.1'],
             dtype='object'
        '''
        print(self.exceptionReportDf.head(8))
        


        self.exceptionReportDf["Routing"] = self.ligne
        self.exceptionReportDf["Material"] = self.exceptionReportDf["Article1"]
        self.exceptionReportDf["Name of sold-to party"] = self.exceptionReportDf["Client1"]
        self.exceptionReportDf["Sales Type"] = self.exceptionReportDf["Présence Prototype"]

        rawFormat = self.exceptionReportDf["FW Format (..x..x..)"].astype(str).str.replace(" mm", "").str.split("x", expand=True)
        if len(rawFormat.columns) == 2:
            rawFormat[2] = 0

        for col in rawFormat.columns:
            for index, value in rawFormat[col].items():
                try:
                    rawFormat.at[index, col] = format_number_eu(value, decimals=2)
                except:
                    rawFormat.at[index, col] = "0"

        self.exceptionReportDf["Length"] = rawFormat[1]
        self.exceptionReportDf["Width"] = rawFormat[2]
        self.exceptionReportDf["Thickness"] = rawFormat[0]



        
        # add a column to exceptionReportDf called "Productivity" and fill it with -1
        self.exceptionReportDf["Productivity"] = -1
        
        # add a column named details taht stores abaques Indexes
        self.exceptionReportDf["Details"] = ""


        subIndex=0
        for index, row in self.exceptionReportDf.iterrows():
            prod, details, level = self.getProductivityForRow(index, row)
            self.exceptionReportDf.at[index, "Productivity"] = round(float(prod), 2)
            self.exceptionReportDf.at[index, "Details"] = str(details) + "#" + str(level)
            
            print(f"Row {subIndex} / {index} / {len(self.exceptionReportDf)} level: {level}, Prod: {round(float(prod), 2)}")
            if subIndex % 20 == 0:
                self.saveCachedProductivities()

            subIndex+=1

        return self.exceptionReportDf

    def getProductivityForRow(self, index, row):
        try:
            articleReport = int(row["Material"])
            clientReport = row["Name of sold-to party"]
            salesTypeReport = row["Sales Type"]
            
            dim1Report = row["Length"]
            dim2Report = row["Width"]
            dim3Report = row["Thickness"]
            dim1Report = round(float(dim1Report), 2)
            dim2Report = round(float(dim2Report), 2)
            dim3Report = round(float(dim3Report), 2)
        except:
            print("Error while parsing row ", index, row["Material"], row["Name of sold-to party"], row["Sales Type"], row["Length"], row["Width"], row["Thickness"])
            return -1, [], -1

        self.curentDetails = []
        self.curentProductivity = -1

        if articleReport in self.articleCachedProductivities:
            return self.articleCachedProductivities[articleReport][0], self.articleCachedProductivities[articleReport][1], self.articleCachedProductivities[articleReport][2]



        self.filterLevel0(articleReport)
        if self.curentProductivity != -1:
            self.articleCachedProductivities[articleReport] = [round(float(self.curentProductivity), 2), self.curentDetails, 0]
            return round(float(self.curentProductivity), 2), self.curentDetails, 0
        

        
        self.filterLevel1(clientReport, salesTypeReport, dim1Report, dim2Report, dim3Report)
        if self.curentProductivity != -1:
            self.articleCachedProductivities[articleReport] = [round(float(self.curentProductivity), 2), self.curentDetails, 1]
            return round(float(self.curentProductivity), 2), self.curentDetails, 1
        
        
        self.filterLevel2(clientReport, salesTypeReport, dim3Report)
        if self.curentProductivity != -1:
            self.articleCachedProductivities[articleReport] = [round(float(self.curentProductivity), 2), self.curentDetails, 2]
            return round(float(self.curentProductivity), 2), self.curentDetails, 2
        

        self.filterLevel3(clientReport, salesTypeReport)
        if self.curentProductivity != -1:
            self.articleCachedProductivities[articleReport] = [round(float(self.curentProductivity), 2), self.curentDetails, 3]
            return round(float(self.curentProductivity), 2), self.curentDetails, 3
        
        
        oldLineReport = row["Routing"]
        self.filterLevel4(oldLineReport, salesTypeReport)
        if self.curentProductivity != -1:
            self.articleCachedProductivities[articleReport] = [round(float(self.curentProductivity), 2), self.curentDetails, 4]
            return round(float(self.curentProductivity), 2), self.curentDetails, 4
        


        print("Absolutly no match for row ", index, articleReport, clientReport, salesTypeReport, dim1Report, dim2Report, dim3Report, oldLineReport)
        self.articleCachedProductivities[articleReport] = [-1, self.curentDetails, -1]
        return -1, self.curentDetails, -1
        
    def isProtoName(self, protoReport):
        if "nan" in str(protoReport)  or protoReport == "" or protoReport == None or  protoReport == "Nan"or  protoReport == "NaN" or protoReport == "NAN":
            return 'FAUX'

        if protoReport != "":
            return 'VRAI'
        return 'FAUX'

    def filterLevel0(self, article):
        potentialProd = []
        self.curentDetails = []

        for index, row in self.abaqueDf.iterrows():
            try:
                if str(int(article)) in str(row["Articles"]):
                    potentialProd.append(row["Prod T/h/OF"])
                    self.curentDetails.append(index)
            except:
                pass

        if len(potentialProd) > 0:
            self.curentProductivity = sum(potentialProd)/len(potentialProd)
        else:
            self.curentProductivity = - 1
    
    def filterLevel1(self, name, proto, dim1, dim2, dim3):
        potentialProd = []
        self.curentDetails = []

        for index, row in self.abaqueDf.iterrows():
            try:
                b1 = str(name) in str(row["Clients"])
                
                if self.isProtoName(proto) == 'VRAI':
                    b2 = int(row["Proto"]) == 1
                else:
                    b2 = int(row["Proto"]) == 0

                
                b3 = dim1 == round(float(row["Épaisseur Nominal"]), 2)
                b4 = dim2 == round(float(row["Largeur"]), 2)
                b5 = dim3 == round(float(row["Longueur"]), 2)

                if b1 and b2 and float(dim1)+float(dim2)+float(dim3) != 0:
                    print(int(b1), int(b2), int(b3), int(b4), int(b5), "Proto match for row ", index, "name: ", name, "proto: ", proto, "dim1: ", dim1, "dim2: ", dim2, "dim3: ", dim3, "comparing to ", row["Clients"], row["Proto"], row["Épaisseur Nominal"], row["Largeur"], row["Longueur"])

                
                if b1 and b2 and b3 and b4 and b5:
                    potentialProd.append(row["Prod T/h/OF"])
                    self.curentDetails.append(index)
            except:
                pass

        if len(potentialProd) > 0:
            self.curentProductivity = sum(potentialProd)/len(potentialProd)
        else:
            self.curentProductivity = -1
           
    def filterLevel2(self, name, proto, dim1):
        potentialProd = []
        self.curentDetails = []

        for index, row in self.abaqueDf.iterrows():
            try:
                b1 = str(name) in str(row["Clients"])
                
                if self.isProtoName(proto) == 'VRAI':
                    b2 = int(row["Proto"]) == 1
                else:
                    b2 = int(row["Proto"]) == 0

                
                b3 = dim1 == round(float(row["Épaisseur Nominal"]), 2)


                if b1 and b2 and b3:
                    potentialProd.append(row["Prod T/h/OF"])
                    self.curentDetails.append(index)
            except:
                pass

        if len(potentialProd) > 0:
            self.curentProductivity = sum(potentialProd)/len(potentialProd)
        else:
            self.curentProductivity = -1
    
    def filterLevel3(self, name, proto):
        potentialProd = []
        self.curentDetails = []

        for index, row in self.abaqueDf.iterrows():
            try:
                b1 = str(name) in str(row["Clients"])
                
                if self.isProtoName(proto) == 'VRAI':
                    b2 = int(row["Proto"]) == 1
                else:
                    b2 = int(row["Proto"]) == 0

                

                if b1 and b2:
                    potentialProd.append(row["Prod T/h/OF"])
                    self.curentDetails.append(index)
            except:
                pass

        if len(potentialProd) > 0:
            self.curentProductivity = sum(potentialProd)/len(potentialProd)
        else:
            self.curentProductivity = -1
        
    def filterLevel4(self, oldLine, proto):
        self.curentDetails = []

        if self.isProtoName(proto) == 'VRAI':
            if oldLine in oldLines:
                self.curentDetails = "Average of line " + str(oldLine) + " for proto articles in abaque"
                self.curentProductivity = avgLineProto[oldLines.index(oldLine)]
                return
        else:
            if oldLine in oldLines:
                self.curentDetails = "Average of line " + str(oldLine) + " for serie articles in abaque"
                self.curentProductivity = avgLineSerie[oldLines.index(oldLine)]
                return
            
        self.curentDetails = "No average line data for " + str(oldLine) + " proto" + str(proto) 
        self.curentProductivity = -1



class previsions():
    def __init__(self):
        self.abaquePath = "Abaque/Abaque.xlsm"
        self.prevision_ligne_Df = self.getPrevision_ligne_Df("R6")
        self.abaqueDf = self.getAbaqueDf()

        mP = MatchingProductivities(self.prevision_ligne_Df, self.abaqueDf, "R6")
        self.mainDF = mP.exceptionReportDf


        processDf()


    def getPrevision_ligne_Df(self, ligne):
        # Load the exception report dataframe from the csv file
        prevision_ligne_Df = pd.read_excel(f"liste_of_ligne/{ligne}.xlsx")
        
        prevision_ligne_Df = prevision_ligne_Df.reset_index(drop=True)
        return prevision_ligne_Df

    def getAbaqueDf(self):
        t1 = time.time()
        print("recuperation of : ", self.abaquePath)
        df = pd.read_excel(self.abaquePath)
        print("recuperation done in : ", time.time() - t1)
        print(df.head())
        print("Abaque length : ", len(df))
        
        return df
    
    def processDf(self):
        self.mainDF["Temps"] = self.mainDF.apply(lambda row: round(float(row[]) / float(row["Productivity"]), 2) if row["Productivity"] != -1 and float(row["Productivity"]) != 0 else -1, axis=1)


prev = previsions()



def format_number_eu(value, decimals=2):
    """
    Format number with:
    - '.' as thousands separator
    - ',' as decimal separator
    - returns '0' if value is None or empty
    """

    if value is None or value == "":
        return "0"

    try:
        number = float(value)
    except (ValueError, TypeError):
        return "0"

    # Format with US style first (1,234.56)
    formatted = f"{number:,.{decimals}f}"

    # Convert to EU style (1.234,56)
    formatted = formatted.replace(",", "X").replace(".", ",").replace("X", ".")

    return formatted