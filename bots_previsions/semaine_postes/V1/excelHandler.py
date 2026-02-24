'''
TODO : 

1. Separate in 3 scripts :
- main.py : contains excelHandler class, and MatchingProductivities class, and all the code to process the exception report and get the productivities
- outputFormatter.py : contains outputFormatter class, and all the code to format the output excel file
- matchingProductivities.py : contains MatchingProductivities class, and all the code to match the productivities

2. Add 'details' option

3. Extend previsions to 8x


V1 - 24/02/2026

Relyes on "Reports/Report.xlsb", "Abaque/Abaque.xlsm", 'OutputTemplate.xlsx'
'''



import os, os.path
import pandas as pd
import time

'Excel formatter dependecy'
from openpyxl import load_workbook

'''

S8
Semaine 16-20

'''
        
#WOIPPY data : 


oldLines =     ["D10", "D10R11", "D14R11", "D20", "D7R10", "FIMI", "FIMIR3B", "L1", "LAS1.", "P3", "R1", "R10", "R11", "R2", "R3B", "R6", "R7", "P3R1"]
newLines =     ["L1", "L1",      "L1",      "L1", "L1",     "L1",  "L1",      "L1", "LASS1,", "P3", "R1", "R1", "R6", "R6", "R6", "R6", "R1", "P3"]
avgLineProto = [4.15, 4.15,      4.15,      4.15, 4.15,     4.15, 4.15,        4.15, 0.39,     2.25, 14,   14,   15.8, 15.8, 15.8, 15.8, 14, 2.25]
avgLineSerie = [6.96, 6.69,      6.69,      6.96, 6.96,     6.96, 6.96,        6.96,  0.7,      2.5, 23.2, 23.2, 39,   39,   39,   39,   23.2,  2.5]

class MatchingProductivities:
    '''
        Takes 2 dataframes as input :
        - exceptionReportDf : dataframe of the exception report, with only Woippy plant data
        - abaqueDf : dataframe of the abaque, with all the data

        Returns the exceptionReportDf with a new column "Productivity" filled with the matched productivity for each row, and a new column "Match Level" filled with the level of the match (0 to 4, or -1 if no match)

        Exemple : 
        
        mP = MatchingProductivities(self.getExceptionReportDf(), self.getAbaqueDf(), self.cacheFile)
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
            
    def __init__(self, exceptionReportDf, abaqueDf, cacheFile="articleCachedProductivities.txt"):
        self.exceptionReportDf = exceptionReportDf
        self.abaqueDf = abaqueDf
        self.cacheFile = cacheFile

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
                    article, prod = line.strip().split(":")
                    articleCachedProductivities[int(article)] = float(prod)
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
                f.write(f"{int(article)}:{round(float(prod), 2)}\n")


    def matchProductivities(self):
        #print("Matching Productivities...")

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

        #print(self.abaqueDf.columns)

        # add a column to exceptionReportDf called "Productivity" and fill it with -1
        self.exceptionReportDf["Productivity"] = -1
        
        subIndex=0
        for index, row in self.exceptionReportDf.iterrows():
            prod, level = self.getProductivityForRow(index, row)
            self.exceptionReportDf.at[index, "Productivity"] = round(float(prod), 2)
            print(f"Row {subIndex} / {index} / {len(self.exceptionReportDf)} level: {level}, Prod: {round(float(prod), 2)}")
            if subIndex % 20 == 0:
                self.saveCachedProductivities()

            subIndex+=1

        return self.exceptionReportDf

    def getProductivityForRow(self, index, row):
        articleReport = int(row["Material"])
        clientReport = row["Name of sold-to party"]
        salesTypeReport = row["Sales Type"]
        dim1Report = row["Length"]
        dim2Report = row["Width"]
        dim3Report = row["Thickness"]

        if articleReport in self.articleCachedProductivities:
            return self.articleCachedProductivities[articleReport], "Cached"

        prod = -1 
        level0 = self.filterLevel0(articleReport)
        if level0 != -1:
            prod = level0
            self.articleCachedProductivities[articleReport] = prod
            return prod, 0
        

        
        level1 = self.filterLevel1(clientReport, salesTypeReport, dim1Report, dim2Report, dim3Report)
        if level1 != -1:
            prod = level1
            self.articleCachedProductivities[articleReport] = prod
            return round(float(prod), 2), 1
        
        level2 = self.filterLevel2(clientReport, salesTypeReport, dim3Report)
        if level2 != -1:
            prod = level2
            self.articleCachedProductivities[articleReport] = prod
            return round(float(prod), 2), 2
    
        level3 = self.filterLevel3(clientReport, salesTypeReport)
        if level3 != -1:
            prod = level3
            self.articleCachedProductivities[articleReport] = prod
            return round(float(prod), 2), 3
        
        oldLineReport = row["Routing"]

        level4 = self.filterLevel4(oldLineReport, salesTypeReport)
        if level4 != -1:
            prod = level4
            self.articleCachedProductivities[articleReport] = prod
            return round(float(prod), 2), 4
        
        print("Absolutly no match for row ", index, articleReport, clientReport, salesTypeReport, dim1Report, dim2Report, dim3Report, oldLineReport)
        self.articleCachedProductivities[articleReport] = -1
        return -1, -1
        

    def isProtoName(self, protoReport):
        if protoReport == "AM Prototype Order" or protoReport == "AM Free Prototype":
            return 'VRAI'
        return 'FAUX'


    def filterLevel0(self, article):
        potentialProd = []

        for index, row in self.abaqueDf.iterrows():
            try:
                if str(int(article)) in str(row["Articles"]):
                    potentialProd.append(row["Prod T/h/OF"])
            except:
                pass

        if len(potentialProd) > 0:
            return sum(potentialProd)/len(potentialProd)
        return -1
    
    def filterLevel1(self, name, proto, dim1, dim2, dim3):
        potentialProd = []

        for index, row in self.abaqueDf.iterrows():
            try:
                if str(name) in str(row["Clients"]) and self.isProtoName(proto) == str(row["Proto"]) and str(dim1) == str(row["Épaisseur Nominal"]) and str(dim2) == str(row["Largeur"]) and str(dim3) == str(row["Longueur"]):
                    potentialProd.append(row["Prod T/h/OF"])
            except:
                pass

        if len(potentialProd) > 0:
            return sum(potentialProd)/len(potentialProd)
        return -1
    
    def filterLevel2(self, name, proto, dim1):
        potentialProd = []

        for index, row in self.abaqueDf.iterrows():
            try:
                if str(name) in str(row["Clients"]) and self.isProtoName(proto) == str(row["Proto"]) and str(dim1) == str(row["Épaisseur Nominal"]):
                    potentialProd.append(row["Prod T/h/OF"])
            except:
                pass

        if len(potentialProd) > 0:
            return sum(potentialProd)/len(potentialProd)
        return -1
    
    def filterLevel3(self, name, proto):
        potentialProd = []

        for index, row in self.abaqueDf.iterrows():
            try:
                if str(name) in str(row["Clients"]) and self.isProtoName(proto) == str(row["Proto"]):
                    potentialProd.append(row["Prod T/h/OF"])
            except:
                pass

        if len(potentialProd) > 0:
            return sum(potentialProd)/len(potentialProd)
        return -1
    
    def filterLevel4(self, oldLine, proto):
        if self.isProtoName(proto) == 'VRAI':
            if oldLine in oldLines:
                return avgLineProto[oldLines.index(oldLine)]
        else:
            if oldLine in oldLines:
                return avgLineSerie[oldLines.index(oldLine)]
            
        return -1

class excelHandler:
    '''
    Opens and get dataframe from Exception_report.xlsb
        => Filters the dataframe to only get Woippy plant data
        => Export dataframe 
    Opens and get dataframe from Abaque.xlsm
        => Export dataframe

    Run MatchingProductivities function to get a productivity for each client in report



    
    '''
    def __init__(self, exceptionReportPath, abaquePath, plantName, cacheFile="articleCachedProductivities.txt"):
        self.exceptionReportPath = exceptionReportPath
        self.abaquePath = abaquePath
        self.plantName = plantName
        self.cacheFile = cacheFile

        
        mP = MatchingProductivities(self.getExceptionReportDf(), self.getAbaqueDf(), self.cacheFile)
        self.newExceptionReportDf = mP.exceptionReportDf
        self.processExceptionReport()

    def getExceptionReportDf(self):
        t1 = time.time()
        sheetName = 8
        print("recuperation of : ", self.exceptionReportPath)
        df = pd.read_excel(self.exceptionReportPath, sheet_name=sheetName)
        print("recuperation done in : ", time.time() - t1)
        newDf = df[df.apply(lambda row: row.astype(str).str.contains(self.plantName, case=False).any(), axis=1)]
        print("filtering done in : ", time.time() - t1)
        print("Time taken: ", time.time() - t1)
        return newDf

    def getAbaqueDf(self):
        t1 = time.time()
        print("recuperation of : ", self.abaquePath)
        df = pd.read_excel(self.abaquePath)
        print("recuperation done in : ", time.time() - t1)
        return df

    def processExceptionReport(self):
        

        intressingFields = ["Routing", "Material", "Productivity", "Forecast W", "Forecast W1", "Forecast W2", "Forecast W3", "Forecast W4", "Forecast W5", "Forecast W6", "Forecast W7", "Forecast W8", "Backlog"]
    
        calculatedFields = ["Forecast W", "Forecast W1", "Forecast W2", "Forecast W3", "Forecast W4", "Forecast W5", "Forecast W6", "Forecast W7", "Forecast W8", "Backlog"]
        for field in calculatedFields:
            newFieldName = field + " (Postes)"
            self.newExceptionReportDf[newFieldName] = self.newExceptionReportDf.apply(lambda row: round(float(row[field]) / float(row["Productivity"]) / 7.5, 2) if row["Productivity"] != -1 and float(row["Productivity"]) != 0 else -1, axis=1)

        # apply newLine to routing column and save it in a new column called "Routing (Postes)"
        self.newExceptionReportDf["New Routing"] = self.newExceptionReportDf.apply(lambda row: newLines[oldLines.index(row["Routing"])] if row["Routing"] in oldLines else row["Routing"], axis=1)
        # apply isProtoName to proto column and save it in a new column called "Is Proto"
        self.newExceptionReportDf["Is Proto"] = self.newExceptionReportDf.apply(lambda row : self.isProtoName(row["Sales Type"]), axis=1)

                                                                                
        newDf = self.newExceptionReportDf[intressingFields + [field + " (Postes)" for field in calculatedFields] + ["New Routing", "Is Proto"]]
        

        print("Number of lines before processing: ", len(newDf))
        newDf = newDf[newDf.apply(lambda row: all(self.isValid(row[field + " (Postes)"]) for field in calculatedFields), axis=1)]
        newDf = newDf[newDf.apply(lambda row: sum(float(row[field]) for field in calculatedFields) != 0, axis=1)]
        print("Number of lines after processing: ", len(newDf))


        # Create a table with each routing, and subdividing it in proto or not, and 1 column Sum of forecast W
        
        #newDf.to_excel("Processed_Exception_report.xlsx", index=False)
        self.newDf = newDf
        
    def get_newDf(self):
        return self.newDf
    
    def isProtoName(self, protoReport):
        if protoReport == "AM Prototype Order" or protoReport == "AM Free Prototype":
            return 'VRAI'
        return 'FAUX'

    def isValid(self, value):
        if value == -1 or value == None or value == "" or str(value).lower() == "nan":
            return False
        return True

class outputFormatter:
    def __init__(self, df):
        self.df = df
        
        self.outputExcel()

    def outputExcel(self, template_path="OutputTemplate.xlsx"):
        # Create a table with each routing, and subdividing it in proto or not, and 1 column Sum of forecast W
        
        wb = load_workbook(template_path)
        ws = wb.active

        start_row = 5  # L1 total starts here
        current_row = start_row



        columnsToSum = ["Backlog (Postes)", "Backlog", "Forecast W (Postes)", "Forecast W", "Forecast W1 (Postes)", "Forecast W1", "Forecast W2 (Postes)", "Forecast W2", "Forecast W3 (Postes)", "Forecast W3", "Forecast W4 (Postes)", "Forecast W4", "Forecast W5 (Postes)", "Forecast W5", "Forecast W6 (Postes)", "Forecast W6", "Forecast W7 (Postes)", "Forecast W7", "Forecast W8 (Postes)", "Forecast W8"]
        summary = self.df.groupby(["New Routing", "Is Proto"]).agg({col: "sum" for col in columnsToSum}).reset_index()

        # Write summary data to Excel
        for index, row in summary.iterrows():
            for i, col in enumerate(columnsToSum):
                ws.cell(row=current_row, column=2 + i).value = round(float(row[col]), 2)
            
            if current_row %3 == 0:
                current_row=current_row+1
            current_row += 1

        # write total in column B for backlog postes, and in column C for backlog, and so on for each forecast W
        for i, col in enumerate(columnsToSum):
            ws.cell(row=current_row-1, column=2 + i).value = round(float(summary[col].sum()), 2)

        # write subTotal in line 4, 7, 10, 13, 16, 19, 22, 25, 28, 31 in column B for backlog postes, and in column C for backlog, and so on for each forecast W
        for i, col in enumerate(columnsToSum):
            ws.cell(row=4, column=2 + i).value = round(float(summary[summary["New Routing"] == "L1"][col].sum()), 2)
            ws.cell(row=7, column=2 + i).value = round(float(summary[summary["New Routing"] == "LASS1,"][col].sum()), 2)
            ws.cell(row=10, column=2 + i).value = round(float(summary[summary["New Routing"] == "P3"][col].sum()), 2)
            ws.cell(row=13, column=2 + i).value = round(float(summary[summary["New Routing"] == "R1"][col].sum()), 2)
            ws.cell(row=16, column=2 + i).value = round(float(summary[summary["New Routing"] == "R6"][col].sum()), 2)



        wb.save(template_path)




eH = excelHandler("Reports/Report.xlsb", "Abaque/Abaque.xlsm", "WOIPPY")
df = eH.get_newDf()
# Read Processed_Exception_report.xlsx and stor it in a dataframe
outputFormatter(df)

