# Date: 2026-03-14
# Author: Dylan H WILDING
# LLMs used : Gemini 3.1 Pro, Claude Sonnet 4.6
# Objective: EPS/PE estimates dashboard, scenario analysis, insider buying
# Project: Holden Valuation Model

import pandas as pd
import numpy as np
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
import yfinance as yf
import datetime
import requests
from bs4 import BeautifulSoup
from pathlib import Path
import argparse

# --- CONFIGURATION ---
# S&P500 TICKERS :
# TICKERS = ["NVDA","AAPL","MSFT","AMZN","GOOGL","GOOG","META","AVGO","TSLA","BRK.B","WMT","LLY","JPM","XOM","V","JNJ","MU","COST","ORCL","MA","NFLX","CVX","ABBV","PLTR","PG","HD","BAC","KO","CAT","AMD","GE","CSCO","MRK","RTX","PM","AMAT","LRCX","UNH","MS","TMUS","MCD","GS","IBM","LIN","WFC","INTC","PEP","GEV","VZ","AXP","AMGN","T","NEE","ABT","KLAC","C","GILD","CRM","DIS","TXN","TJX","TMO","ANET","ISRG","BA","APH","SCHW","DE","APP","PFE","UBER","ADI","HON","COP","LMT","WELL","UNP","BLK","QCOM","ETN","PANW","BKNG","LOW","DHR","SYK","CB","SPGI","PLD","INTU","ACN","BMY","PGR","VRTX","NEM","HCA","NOW","MCK","MO","SBUX","PH","CRWD","MDT","COF","CME","GLW","SO","CEG","CMCSA","NOC","DUK","BSX","ADBE","DELL","SNDK","CVS","WM","EQIX","GD","HWM","TT","WDC","ICE","WMB","STX","AMT","ADP","PWR","BX","MRSH","MAR","FDX","UPS","PNC","FCX","BK","NKE","JCI","MMM","USB","SHW","CDNS","SNPS","REGN","MSI","CTAS","ECL","ITW","ORLY","KKR","MCO","ABNB","MNST","EMR","KMI","CMI","RCL","CSX","AEP","EOG","CL","CI","MDLZ","DASH","RSG","PSX","VLO","AON","TDG","COR","WBD","SLB","LHX","HLT","CRH","ROST","MPC","HOOD","GM","TRV","NSC","ELV","APD","SRE","FTNT","DLR","SPG","PCAR","APO","O","AZO","TEL","OXY","AFL","D","TFC","OKE","BKR","VST","ALL","AJG","CTVA","TGT","ADSK","PSA","FAST","MPWR","COIN","TRGP","FANG","XEL","CAH","EXC","GWW","EA","AME","ZTS","NDAQ","NXPI","KEYS","FIX","ETR","CIEN","EW","F","CARR","URI","KR","IDXX","BDX","GRMN","TER","YUM","MET","DDOG","HSY","CVNA","CMG","PEG","ED","PYPL","AIG","VTR","SYY","EBAY","DHI","ROK","WAB","AMP","MSCI","EQT","AXON","PCG","CBRE","FITB","TTWO","DAL","WEC","CCI","ODFL","KDP","NUE","HIG","TPL","LYV","ROP","LVS","XYZ","MLM","WDAY","VMC","ADM","STT","RMD","KVUE","MCHP","ACGL","PAYX","CCL","CPRT","KMB","NRG","OTIS","GEHC","IR","PRU","EL","EME","IRM","A","ATO","DTE","AEE","HBAN","VICI","FISV","CBOE","EXR","TDY","FE","IBKR","MTB","XYL","DG","PPL","CTSH","CNP","TPR","RJF","DVN","HPE","HAL","WAT","UAL","EXPE","VRSK","IQV","CHTR","EIX","DOV","ES","WTW","AWK","FICO","KHC","BIIB","ROL","JBL","DOW","STZ","WRB","EXE","FIS","CINF","NTRS","TSCO","HUBB","DXCM","STLD","CTRA","OMC","BG","MTD","CMS","ULTA","AVB","CFG","LEN","DRI","LYB","BRO","CHD","ON","PHM","Q","ARES","PPG","NI","EQR","VLTO","L","SYF","EFX","LDOS","DGX","VRSN","LH","CPAY","RF","WSM","DLTR","TSN","STE","FSLR","GIS","SW","BR","MRNA","KEY","CHRW","RL","CF","SBAC","HUM","IP","NTAP","TROW","GPN","PKG","SNA","LUV","EXPD","EVRG","JBHT","AMCR","LNT","ALB","PFG","PTC","LULU","SMCI","INCY","DD","CSGP","ZBH","NVR","IFF","HPQ","WST","CNC","WY","HOLX","LII","FTV","BALL","FFIV","HII","ESS","TXT","MKC","AKAM","VTRS","TRMB","PODD","KIM","INVH","TKO","J","TYL","APTV","CDW","NDSN","MAA","GPC","PNR","DECK","REG","IEX","COO","DPZ","CLX","HAS","AVY","BBY","TTD","EG","ERIE","HST","BEN","MAS","GEN","ALLE","HRL","PNW","JKHY","APA","DOC","GNRC","UHS","IT","ALGN","FOX","UDR","SOLV","FOXA","SJM","SWK","GL","AIZ","GDDY","PSKY","BF.B","WYNN","CPT","IVZ","AES","DVA","ZBRA","BLDR","RVTY","MGM","MOS","AOS","FRT","BAX","NWSA","HSIC","NCLH","ARE","BXP","SWKS","TAP","TECH","CAG","MOH","CRL","FDS","POOL","EPAM","MTCH","PAYC","CPB","LW","NWS"]

# NASDAQ100 TICKERS : 
# TICKERS = ["NVDA","AAPL","MSFT","AMZN","GOOGL","GOOG","AVGO","META","TSLA","WMT","ASML","COST","NFLX","MU","PLTR","AMD","CSCO","AMAT","LRCX","TMUS","LIN","PEP","INTC","AMGN","KLAC","TXN","GILD","ISRG","ADI","SHOP","ARM","HON","PDD","QCOM","BKNG","APP","PANW","INTU","VRTX","CEG","CMCSA","SBUX","ADBE","CRWD","WDC","MAR","ADP","MELI","STX","REGN","ORLY","MRVL","CDNS","ABNB","CSX","SNPS","MDLZ","AEP","MNST","CTAS","WBD","ROST","DASH","BKR","PCAR","FTNT","FANG","FAST","EA","EXC","XEL","ADSK","MPWR","NXPI","IDXX","FER","MSTR","ALNY","DDOG","PYPL","CCEP","TRI","ODFL","KDP","ROP","TTWO","INSM","AXON","WDAY","PAYX","MCHP","GEHC","CPRT","CTSH","CHTR","KHC","VRSK","DXCM","ZS","TEAM","CSGP"]

# RUSELL 3000 TICKERS : 
# TICKERS = ['A', 'AA', 'AAL', 'AAMI', 'AAOI', 'AAON', 'AAP', 'AAPL', 'AAT', 'ABAT', 'ABBV', 'ABCB', 'ABG', 'ABM', 'ABNB', 'ABR', 'ABSI', 'ABT', 'ABUS', 'ACA', 'ACAD', 'ACCO', 'ACDC', 'ACEL', 'ACGL', 'ACH', 'ACHC', 'ACHR', 'ACI', 'ACIC', 'ACIW', 'ACLS', 'ACLX', 'ACM', 'ACMR', 'ACN', 'ACNB', 'ACRE', 'ACT', 'ACTG', 'ACVA', 'ADAM', 'ADBE', 'ADC', 'ADCT', 'ADEA', 'ADI', 'ADM', 'ADMA', 'ADNT', 'ADP', 'ADPT', 'ADRO', 'ADSK', 'ADT', 'ADTN', 'ADUS', 'ADV', 'AEBI', 'AEE', 'AEHR', 'AEIS', 'AEO', 'AEP', 'AES', 'AESI', 'AEVA', 'AFG', 'AFL', 'AFRM', 'AGCO', 'AGIO', 'AGL', 'AGM', 'AGNC', 'AGO', 'AGX', 'AGYS', 'AHCO', 'AHR', 'AHRT', 'AI', 'AIG', 'AIN', 'AIOT', 'AIP', 'AIR', 'AIT', 'AIZ', 'AJG', 'AKAM', 'AKBA', 'AKE', 'AKR', 'AKTS', 'AL', 'ALAB', 'ALB', 'ALDX', 'ALEC', 'ALG', 'ALGM', 'ALGN', 'ALGT', 'ALH', 'ALHC', 'ALIT', 'ALK', 'ALKS', 'ALKT', 'ALL', 'ALLE', 'ALLO', 'ALLY', 'ALMS', 'ALNT', 'ALNY', 'ALRM', 'ALRS', 'ALSN', 'ALT', 'ALTG', 'ALX', 'AM', 'AMAL', 'AMAT', 'AMBA', 'AMBP', 'AMBQ', 'AMC', 'AMCR', 'AMCX', 'AMD', 'AME', 'AMG', 'AMGN', 'AMH', 'AMKR', 'AMLX', 'AMN', 'AMP', 'AMPH', 'AMPL', 'AMPX', 'AMR', 'AMRC', 'AMRX', 'AMSC', 'AMSF', 'AMT', 'AMTB', 'AMTM', 'AMWD', 'AMZN', 'AN', 'ANAB', 'ANDE', 'ANET', 'ANF', 'ANGI', 'ANGO', 'ANIK', 'ANIP', 'ANNX', 'AON', 'AORT', 'AOS', 'AOSL', 'APA', 'APAM', 'APD', 'APEI', 'APG', 'APGE', 'APH', 'APLD', 'APLE', 'APLS', 'APO', 'APOG', 'APP', 'APPF', 'APPN', 'APPS', 'APTV', 'AQST', 'AR', 'ARAY', 'ARCB', 'ARCT', 'ARDX', 'ARE', 'ARES', 'ARHS', 'ARI', 'ARKO', 'ARLO', 'ARMK', 'AROC', 'AROW', 'ARQT', 'ARR', 'ARRY', 'ARVN', 'ARW', 'ARWR', 'AS', 'ASAN', 'ASB', 'ASC', 'ASGN', 'ASH', 'ASIX', 'ASLE', 'ASO', 'ASPI', 'ASPN', 'ASST', 'ASTE', 'ASTH', 'ASTS', 'ASUR', 'ATEC', 'ATEN', 'ATEX', 'ATI', 'ATKR', 'ATMU', 'ATNI', 'ATO', 'ATR', 'ATRC', 'ATRO', 'ATYR', 'AU', 'AUB', 'AUPH', 'AUR', 'AURA', 'AVA', 'AVAH', 'AVAV', 'AVB', 'AVBP', 'AVD', 'AVGO', 'AVIR', 'AVNS', 'AVNT', 'AVNW', 'AVO', 'AVPT', 'AVT', 'AVTR', 'AVXL', 'AVY', 'AWI', 'AWK', 'AWR', 'AX', 'AXGN', 'AXON', 'AXP', 'AXS', 'AXSM', 'AXTA', 'AYI', 'AZO', 'AZTA', 'AZZ', 'BA', 'BAC', 'BAH', 'BALL', 'BAM', 'BANC', 'BAND', 'BANF', 'BANR', 'BATRA', 'BATRK', 'BAX', 'BBAI', 'BBBY', 'BBIO', 'BBNX', 'BBSI', 'BBT', 'BBUC', 'BBW', 'BBWI', 'BBY', 'BC', 'BCAL', 'BCAX', 'BCBP', 'BCC', 'BCML', 'BCO', 'BCPC', 'BCRX', 'BDC', 'BDN', 'BDX', 'BE', 'BEAM', 'BELFB', 'BEN', 'BEPC', 'BETA', 'BETR', 'BFA', 'BFAM', 'BFB', 'BFC', 'BFH', 'BFLY', 'BFS', 'BFST', 'BG', 'BGC', 'BGS', 'BHB', 'BHE', 'BHF', 'BHR', 'BHRB', 'BHVN', 'BIIB', 'BILL', 'BIO', 'BIOA', 'BIPC', 'BIRK', 'BJ', 'BJRI', 'BK', 'BKD', 'BKE', 'BKH', 'BKNG', 'BKR', 'BKSY', 'BKU', 'BKV', 'BL', 'BLBD', 'BLD', 'BLDR', 'BLFS', 'BLFY', 'BLK', 'BLKB', 'BLMN', 'BLND', 'BLSH', 'BLX', 'BMBL', 'BMI', 'BMRC', 'BMRN', 'BMY', 'BNL', 'BOC', 'BOH', 'BOKF', 'BOOM', 'BOOT', 'BORR', 'BOW', 'BOX', 'BPOP', 'BR', 'BRBR', 'BRCC', 'BRKB', 'BRKR', 'BRO', 'BROS', 'BRSL', 'BRSP', 'BRX', 'BRZE', 'BSRR', 'BSX', 'BSY', 'BTBT', 'BTDR', 'BTSG', 'BTU', 'BULL', 'BUR', 'BURL', 'BUSE', 'BV', 'BVS', 'BWA', 'BWB', 'BWIN', 'BWMN', 'BWXT', 'BX', 'BXC', 'BXMT', 'BXP', 'BY', 'BYD', 'BYND', 'BYRN', 'BZH', 'C', 'CABO', 'CAC', 'CACC', 'CACI', 'CAG', 'CAH', 'CAI', 'CAKE', 'CAL', 'CALM', 'CALX', 'CALY', 'CAPR', 'CAR', 'CARE', 'CARG', 'CARR', 'CARS', 'CART', 'CASH', 'CASS', 'CASY', 'CAT', 'CATX', 'CATY', 'CAVA', 'CB', 'CBAN', 'CBL', 'CBLL', 'CBOE', 'CBRE', 'CBRL', 'CBSH', 'CBT', 'CBU', 'CBZ', 'CC', 'CCB', 'CCBG', 'CCC', 'CCI', 'CCK', 'CCL', 'CCNE', 'CCOI', 'CCRN', 'CCS', 'CCSI', 'CDE', 'CDNA', 'CDNS', 'CDP', 'CDRE', 'CDW', 'CDXS', 'CE', 'CECO', 'CEG', 'CELC', 'CELH', 'CENT', 'CENTA', 'CENX', 'CERS', 'CERT', 'CEVA', 'CF', 'CFFN', 'CFG', 'CFR', 'CG', 'CGEM', 'CGNX', 'CGON', 'CHCO', 'CHCT', 'CHD', 'CHDN', 'CHE', 'CHEF', 'CHH', 'CHRD', 'CHRS', 'CHRW', 'CHTR', 'CHWY', 'CI', 'CIEN', 'CIFR', 'CIM', 'CINF', 'CIVB', 'CL', 'CLB', 'CLBK', 'CLDT', 'CLDX', 'CLF', 'CLFD', 'CLH', 'CLMB', 'CLMT', 'CLNE', 'CLOV', 'CLPT', 'CLSK', 'CLVT', 'CLW', 'CLX', 'CMC', 'CMCL', 'CMCO', 'CMCSA', 'CMDB', 'CME', 'CMG', 'CMI', 'CMP', 'CMPR', 'CMPX', 'CMRC', 'CMRE', 'CMS', 'CMT', 'CMTG', 'CNA', 'CNC', 'CNDT', 'CNH', 'CNK', 'CNM', 'CNMD', 'CNNE', 'CNO', 'CNOB', 'CNP', 'CNR', 'CNS', 'CNX', 'CNXC', 'CNXN', 'COCO', 'CODI', 'COF', 'COFS', 'COGT', 'COHR', 'COHU', 'COIN', 'COKE', 'COLB', 'COLD', 'COLL', 'COLM', 'COMP', 'CON', 'COO', 'COP', 'COR', 'CORT', 'CORZ', 'COST', 'COTY', 'COUR', 'CPAY', 'CPB', 'CPF', 'CPK', 'CPNG', 'CPRI', 'CPRT', 'CPRX', 'CPS', 'CPT', 'CR', 'CRAI', 'CRC', 'CRCL', 'CRCT', 'CRDO', 'CRGY', 'CRH', 'CRI', 'CRK', 'CRL', 'CRM', 'CRMD', 'CRML', 'CRMT', 'CRNC', 'CRNX', 'CROX', 'CRS', 'CRSP', 'CRSR', 'CRUS', 'CRVL', 'CRVS', 'CRWD', 'CSCO', 'CSGP', 'CSGS', 'CSL', 'CSR', 'CSTL', 'CSTM', 'CSV', 'CSW', 'CSX', 'CTAS', 'CTBI', 'CTKB', 'CTLP', 'CTO', 'CTOS', 'CTRA', 'CTRE', 'CTRI', 'CTS', 'CTSH', 'CTVA', 'CUBE', 'CUBI', 'CURB', 'CUZ', 'CVBF', 'CVCO', 'CVGW', 'CVI', 'CVLG', 'CVLT', 'CVNA', 'CVRX', 'CVS', 'CVSA', 'CVX', 'CW', 'CWAN', 'CWBC', 'CWCO', 'CWEN', 'CWENA', 'CWH', 'CWK', 'CWST', 'CWT', 'CXM', 'CXT', 'CXW', 'CYH', 'CYRX', 'CYTK', 'CZFS', 'CZNC', 'CZR', 'D', 'DAKT', 'DAL', 'DAN', 'DAR', 'DASH', 'DAVE', 'DAWN', 'DBD', 'DBI', 'DBRG', 'DBX', 'DC', 'DCGO', 'DCH', 'DCI', 'DCO', 'DCOM', 'DD', 'DDD', 'DDOG', 'DDS', 'DE', 'DEA', 'DEC', 'DECK', 'DEI', 'DELL', 'DFH', 'DFIN', 'DFTX', 'DG', 'DGICA', 'DGII', 'DGX', 'DH', 'DHC', 'DHI', 'DHIL', 'DHR', 'DHT', 'DIN', 'DINO', 'DIOD', 'DIS', 'DJCO', 'DJT', 'DK', 'DKNG', 'DKS', 'DLB', 'DLR', 'DLTR', 'DLX', 'DMRC', 'DNLI', 'DNOW', 'DNTH', 'DNUT', 'DOC', 'DOCN', 'DOCS', 'DOCU', 'DOLE', 'DOMO', 'DORM', 'DOV', 'DOW', 'DOX', 'DPZ', 'DRH', 'DRI', 'DRS', 'DRUG', 'DRVN', 'DSGN', 'DSGR', 'DSP', 'DT', 'DTE', 'DTM', 'DUK', 'DUOL', 'DV', 'DVA', 'DVN', 'DX', 'DXC', 'DXCM', 'DXPE', 'DY', 'DYN', 'EA', 'EAT', 'EBAY', 'EBC', 'EBF', 'EBS', 'ECG', 'ECL', 'ECPG', 'ECVT', 'ED', 'EDIT', 'EE', 'EEFT', 'EFC', 'EFSC', 'EFX', 'EG', 'EGBN', 'EGHT', 'EGP', 'EGY', 'EHAB', 'EHC', 'EIG', 'EIX', 'EL', 'ELAN', 'ELF', 'ELS', 'ELV', 'ELVN', 'EMBC', 'EME', 'EMN', 'EMR', 'ENOV', 'ENPH', 'ENR', 'ENS', 'ENSG', 'ENTA', 'ENTG', 'ENVA', 'ENVX', 'EOG', 'EOLS', 'EOSE', 'EPAC', 'EPAM', 'EPC', 'EPM', 'EPR', 'EPRT', 'EQBK', 'EQH', 'EQIX', 'EQR', 'EQT', 'ERAS', 'ERII', 'ES', 'ESAB', 'ESE', 'ESI', 'ESNT', 'ESPR', 'ESQ', 'ESRT', 'ESS', 'ESTC', 'ETD', 'ETN', 'ETON', 'ETR', 'ETSY', 'EU', 'EVC', 'EVCM', 'EVER', 'EVGO', 'EVH', 'EVLV', 'EVR', 'EVRG', 'EVTC', 'EW', 'EWBC', 'EWCZ', 'EWTX', 'EXC', 'EXE', 'EXEL', 'EXLS', 'EXP', 'EXPD', 'EXPE', 'EXPI', 'EXPO', 'EXR', 'EXTR', 'EYE', 'EYPT', 'F', 'FA', 'FAF', 'FANG', 'FAST', 'FATE', 'FBIN', 'FBIZ', 'FBK', 'FBNC', 'FBP', 'FBRT', 'FC', 'FCBC', 'FCF', 'FCFS', 'FCN', 'FCNCA', 'FCPT', 'FCX', 'FDBC', 'FDMT', 'FDP', 'FDS', 'FDX', 'FE', 'FELE', 'FERG', 'FET', 'FFBC', 'FFIC', 'FFIN', 'FFIV', 'FFWM', 'FG', 'FHB', 'FHN', 'FIBK', 'FICO', 'FIGR', 'FIGS', 'FIHL', 'FIP', 'FIS', 'FISI', 'FISV', 'FITB', 'FIVE', 'FIVN', 'FIX', 'FIZZ', 'FLEX', 'FLG', 'FLGT', 'FLNC', 'FLNG', 'FLO', 'FLOC', 'FLR', 'FLS', 'FLUT', 'FLWS', 'FLY', 'FLYW', 'FMAO', 'FMBH', 'FMC', 'FMNB', 'FN', 'FNB', 'FND', 'FNF', 'FNKO', 'FNLC', 'FOLD', 'FOR', 'FORM', 'FORR', 'FOUR', 'FOX', 'FOXA', 'FOXF', 'FPI', 'FR', 'FRBA', 'FRHC', 'FRME', 'FRPH', 'FRPT', 'FRSH', 'FRST', 'FRT', 'FSBC', 'FSBW', 'FSLR', 'FSLY', 'FSS', 'FSUN', 'FTAI', 'FTDR', 'FTI', 'FTNT', 'FTRE', 'FTV', 'FUBO', 'FUL', 'FULC', 'FULT', 'FUN', 'FWONA', 'FWONK', 'FWRD', 'FWRG', 'G', 'GABC', 'GAP', 'GATX', 'GBCI', 'GBTG', 'GBX', 'GCMG', 'GCO', 'GCT', 'GD', 'GDDY', 'GDEN', 'GDOT', 'GDYN', 'GE', 'GEF', 'GEFB', 'GEHC', 'GEN', 'GENI', 'GEO', 'GERN', 'GETY', 'GEV', 'GEVO', 'GFF', 'GFS', 'GGG', 'GH', 'GHC', 'GHM', 'GIC', 'GIII', 'GILD', 'GIS', 'GKOS', 'GL', 'GLDD', 'GLIBA', 'GLIBK', 'GLNG', 'GLOB', 'GLPI', 'GLRE', 'GLUE', 'GLW', 'GM', 'GME', 'GMED', 'GNE', 'GNK', 'GNL', 'GNRC', 'GNTX', 'GNW', 'GO', 'GOGO', 'GOLD', 'GOLF', 'GOOD', 'GOOG', 'GOOGL', 'GOSS', 'GPC', 'GPGI', 'GPI', 'GPK', 'GPN', 'GPOR', 'GPRE', 'GRAL', 'GRBK', 'GRC', 'GRDN', 'GRMN', 'GRND', 'GRNT', 'GRPN', 'GS', 'GSAT', 'GSBC', 'GSHD', 'GT', 'GTES', 'GTLB', 'GTLS', 'GTM', 'GTN', 'GTX', 'GTXI', 'GTY', 'GVA', 'GWRE', 'GWW', 'GXO', 'H', 'HAE', 'HAFC', 'HAIN', 'HAL', 'HALO', 'HAS', 'HASI', 'HAYW', 'HBAN', 'HBCP', 'HBNC', 'HBT', 'HCA', 'HCAT', 'HCC', 'HCI', 'HCKT', 'HCSG', 'HD', 'HDSN', 'HE', 'HEI', 'HEIA', 'HELE', 'HFWA', 'HG', 'HGV', 'HHH', 'HIFS', 'HIG', 'HII', 'HIMS', 'HIPO', 'HIW', 'HL', 'HLF', 'HLI', 'HLIO', 'HLIT', 'HLMN', 'HLNE', 'HLT', 'HLX', 'HMN', 'HNI', 'HNRG', 'HNST', 'HOG', 'HOLX', 'HOMB', 'HON', 'HOOD', 'HOPE', 'HOV', 'HP', 'HPE', 'HPP', 'HPQ', 'HQY', 'HR', 'HRB', 'HRI', 'HRL', 'HRMY', 'HROW', 'HRTG', 'HRTX', 'HSIC', 'HST', 'HSTM', 'HSY', 'HTB', 'HTBK', 'HTFL', 'HTH', 'HTLD', 'HTO', 'HTZ', 'HUBB', 'HUBG', 'HUBS', 'HUM', 'HUMA', 'HUN', 'HURN', 'HUT', 'HVT', 'HWC', 'HWKN', 'HWM', 'HXL', 'HY', 'HYLN', 'HZO', 'IAC', 'IART', 'IBCP', 'IBKR', 'IBM', 'IBOC', 'IBP', 'IBRX', 'IBTA', 'ICE', 'ICFI', 'ICHR', 'ICUI', 'IDA', 'IDCC', 'IDR', 'IDT', 'IDXX', 'IDYA', 'IE', 'IESC', 'IEX', 'IFF', 'IHRT', 'IIIN', 'IIIV', 'IIPR', 'ILMN', 'ILPT', 'IMAX', 'IMKTA', 'IMMR', 'IMNM', 'IMVT', 'IMXI', 'INBK', 'INBX', 'INCY', 'INDB', 'INDI', 'INDV', 'INGM', 'INGN', 'INGR', 'INH', 'INN', 'INOD', 'INR', 'INSE', 'INSM', 'INSP', 'INSW', 'INTA', 'INTC', 'INTU', 'INVA', 'INVH', 'INVX', 'IONQ', 'IONS', 'IOSP', 'IOT', 'IOVA', 'IP', 'IPAR', 'IPGP', 'IPI', 'IQV', 'IR', 'IRDM', 'IRM', 'IRMD', 'IRON', 'IRT', 'IRTC', 'IRWD', 'ISRG', 'IT', 'ITGR', 'ITIC', 'ITRI', 'ITT', 'ITW', 'IVR', 'IVT', 'IVZ', 'J', 'JACK', 'JAKK', 'JANX', 'JAZZ', 'JBGS', 'JBHT', 'JBI', 'JBIO', 'JBL', 'JBLU', 'JBSS', 'JBTM', 'JCI', 'JEF', 'JELD', 'JHG', 'JHX', 'JJSF', 'JKHY', 'JLL', 'JMSB', 'JNJ', 'JOBY', 'JOE', 'JOUT', 'JPM', 'JRVR', 'JXN', 'KAI', 'KALU', 'KALV', 'KBH', 'KBR', 'KD', 'KDP', 'KE', 'KELYA', 'KEX', 'KEY', 'KEYS', 'KFRC', 'KFY', 'KGS', 'KHC', 'KIDS', 'KIM', 'KKR', 'KLAC', 'KLC', 'KLIC', 'KMB', 'KMI', 'KMPR', 'KMT', 'KMTS', 'KMX', 'KN', 'KNF', 'KNSL', 'KNTK', 'KNX', 'KO', 'KOD', 'KODK', 'KOP', 'KOPN', 'KOS', 'KR', 'KRC', 'KREF', 'KRG', 'KRMN', 'KRNY', 'KROS', 'KRRO', 'KRUS', 'KRYS', 'KSS', 'KTB', 'KTOS', 'KURA', 'KVUE', 'KW', 'KWR', 'KYMR', 'L', 'LAB', 'LAD', 'LADR', 'LAMR', 'LAND', 'LASR', 'LAUR', 'LAW', 'LAZ', 'LBRDA', 'LBRDK', 'LBRT', 'LBTYA', 'LBTYK', 'LC', 'LCID', 'LCII', 'LDOS', 'LEA', 'LECO', 'LEG', 'LEN', 'LENB', 'LENZ', 'LEU', 'LFST', 'LFUS', 'LGIH', 'LGN', 'LGND', 'LH', 'LHX', 'LIF', 'LII', 'LILA', 'LILAK', 'LIN', 'LINC', 'LIND', 'LINE', 'LION', 'LITE', 'LIVN', 'LKFN', 'LKQ', 'LLY', 'LLYVA', 'LLYVK', 'LMAT', 'LMB', 'LMND', 'LMNR', 'LMT', 'LNC', 'LNG', 'LNN', 'LNT', 'LNTH', 'LOAR', 'LOB', 'LOCO', 'LOPE', 'LOVE', 'LOW', 'LPG', 'LPLA', 'LPRO', 'LPX', 'LQDA', 'LQDT', 'LRCX', 'LRMR', 'LRN', 'LSCC', 'LSTR', 'LTC', 'LTH', 'LULU', 'LUMN', 'LUNG', 'LUNR', 'LUV', 'LVS', 'LW', 'LXEO', 'LXFR', 'LXP', 'LXU', 'LYB', 'LYFT', 'LYTS', 'LYV', 'LZ', 'LZB', 'M', 'MA', 'MAA', 'MAC', 'MAMA', 'MAN', 'MANH', 'MAR', 'MARA', 'MAS', 'MASI', 'MASS', 'MAT', 'MATV', 'MATW', 'MATX', 'MAX', 'MAZE', 'MBC', 'MBI', 'MBIN', 'MBUU', 'MBWM', 'MBX', 'MC', 'MCB', 'MCBS', 'MCD', 'MCFT', 'MCHB', 'MCHP', 'MCK', 'MCO', 'MCRI', 'MCS', 'MCW', 'MCY', 'MD', 'MDB', 'MDGL', 'MDLN', 'MDLZ', 'MDT', 'MDU', 'MDXG', 'MED', 'MEDP', 'MEG', 'MEI', 'MET', 'META', 'METC', 'MFA', 'MGEE', 'MGM', 'MGNI', 'MGPI', 'MGRC', 'MGTX', 'MGY', 'MHK', 'MHO', 'MIAX', 'MIDD', 'MIR', 'MIRM', 'MITK', 'MKC', 'MKL', 'MKSI', 'MKTX', 'MLAB', 'MLI', 'MLKN', 'MLM', 'MLR', 'MLYS', 'MMI', 'MMM', 'MMS', 'MMSI', 'MNKD', 'MNRO', 'MNST', 'MNTK', 'MO', 'MOD', 'MOGA', 'MOH', 'MORN', 'MOS', 'MOV', 'MP', 'MPB', 'MPC', 'MPLT', 'MPT', 'MPWR', 'MQ', 'MRCY', 'MRK', 'MRNA', 'MRP', 'MRSH', 'MRTN', 'MRVI', 'MRVL', 'MRX', 'MS', 'MSA', 'MSBI', 'MSCI', 'MSEX', 'MSFT', 'MSGE', 'MSGS', 'MSI', 'MSM', 'MSTR', 'MTB', 'MTCH', 'MTD', 'MTDR', 'MTG', 'MTH', 'MTN', 'MTRN', 'MTSI', 'MTUS', 'MTW', 'MTX', 'MTZ', 'MU', 'MUR', 'MUSA', 'MVBF', 'MVIS', 'MVST', 'MWA', 'MXCT', 'MXL', 'MYE', 'MYGN', 'MYPS', 'MYRG', 'MZTI', 'NABL', 'NAGE', 'NAT', 'NATL', 'NAVI', 'NAVN', 'NB', 'NBBK', 'NBHC', 'NBIX', 'NBN', 'NBR', 'NBTB', 'NCLH', 'NCMI', 'NCNO', 'NDAQ', 'NDSN', 'NE', 'NECB', 'NEE', 'NEM', 'NEO', 'NEOG', 'NESR', 'NET', 'NEU', 'NEWT', 'NEXT', 'NFBK', 'NFE', 'NFG', 'NFLX', 'NG', 'NGNE', 'NGVC', 'NGVT', 'NHC', 'NHI', 'NI', 'NIC', 'NIQ', 'NJR', 'NKE', 'NKTX', 'NLOP', 'NLY', 'NMIH', 'NMRK', 'NN', 'NNE', 'NNI', 'NNN', 'NNOX', 'NOC', 'NOG', 'NOV', 'NOVT', 'NOW', 'NPCE', 'NPK', 'NPKI', 'NPO', 'NRC', 'NRDS', 'NRG', 'NRIM', 'NRIX', 'NSA', 'NSC', 'NSIT', 'NSP', 'NSSC', 'NTAP', 'NTB', 'NTCT', 'NTGR', 'NTLA', 'NTNX', 'NTRA', 'NTRS', 'NTST', 'NU', 'NUE', 'NUS', 'NUTX', 'NUVB', 'NUVL', 'NVAX', 'NVCR', 'NVDA', 'NVEC', 'NVR', 'NVRI', 'NVST', 'NVT', 'NVTS', 'NWBI', 'NWE', 'NWL', 'NWN', 'NWPX', 'NWS', 'NWSA', 'NX', 'NXDR', 'NXDT', 'NXRT', 'NXST', 'NXT', 'NYT', 'O', 'OABI', 'OBK', 'OBT', 'OC', 'OCFC', 'OCUL', 'ODC', 'ODFL', 'OEC', 'OFG', 'OFIX', 'OFLX', 'OGE', 'OGN', 'OGS', 'OHI', 'OI', 'OII', 'OIS', 'OKE', 'OKLO', 'OKTA', 'OLED', 'OLLI', 'OLMA', 'OLN', 'OLP', 'OLPX', 'OMC', 'OMCL', 'OMDA', 'OMER', 'OMF', 'ON', 'ONB', 'ONEW', 'ONIT', 'ONON', 'ONTF', 'ONTO', 'OOMA', 'OPCH', 'OPK', 'OPLN', 'OPRX', 'OPTU', 'ORA', 'ORC', 'ORCL', 'ORGO', 'ORI', 'ORIC', 'ORKA', 'ORLY', 'ORRF', 'OSBC', 'OSCR', 'OSG', 'OSIS', 'OSK', 'OSPN', 'OSUR', 'OSW', 'OTIS', 'OTTR', 'OUST', 'OUT', 'OVV', 'OWL', 'OXM', 'OXY', 'OZK', 'P5N994', 'PACB', 'PACK', 'PACS', 'PAG', 'PAGS', 'PAHC', 'PANL', 'PANW', 'PAR', 'PARR', 'PATH', 'PATK', 'PAX', 'PAYC', 'PAYO', 'PAYS', 'PAYX', 'PB', 'PBF', 'PBH', 'PBI', 'PCAR', 'PCG', 'PCOR', 'PCRX', 'PCT', 'PCTY', 'PCVX', 'PD', 'PDFS', 'PDM', 'PEB', 'PEBO', 'PECO', 'PEG', 'PEGA', 'PEN', 'PENG', 'PENN', 'PEP', 'PFBC', 'PFE', 'PFG', 'PFGC', 'PFIS', 'PFS', 'PFSI', 'PG', 'PGC', 'PGEN', 'PGNY', 'PGR', 'PGY', 'PH', 'PHAT', 'PHIN', 'PHM', 'PHR', 'PI', 'PII', 'PINS', 'PIPR', 'PJT', 'PK', 'PKE', 'PKG', 'PKST', 'PL', 'PLAB', 'PLAY', 'PLD', 'PLMR', 'PLNT', 'PLOW', 'PLPC', 'PLSE', 'PLTK', 'PLTR', 'PLUG', 'PLUS', 'PLXS', 'PM', 'PMT', 'PNC', 'PNFP', 'PNR', 'PNTG', 'PNW', 'PODD', 'POOL', 'POR', 'POST', 'POWI', 'POWL', 'POWW', 'PPC', 'PPG', 'PPL', 'PPTA', 'PR', 'PRA', 'PRAA', 'PRAX', 'PRCH', 'PRCT', 'PRDO', 'PRG', 'PRGO', 'PRGS', 'PRI', 'PRIM', 'PRK', 'PRKS', 'PRLB', 'PRM', 'PRMB', 'PRME', 'PRSU', 'PRTA', 'PRTH', 'PRU', 'PRVA', 'PSA', 'PSFE', 'PSIX', 'PSMT', 'PSN', 'PSTG', 'PSTL', 'PSX', 'PTC', 'PTCT', 'PTEN', 'PTGX', 'PTLO', 'PTON', 'PUBM', 'PUMP', 'PVH', 'PVLA', 'PWP', 'PWR', 'PYPL', 'PZZA', 'Q', 'QBTS', 'QCOM', 'QCRH', 'QDEL', 'QGEN', 'QLYS', 'QNST', 'QRVO', 'QS', 'QSI', 'QSR', 'QTRX', 'QTWO', 'QUBT', 'QXO', 'R', 'RAL', 'RAMP', 'RAPP', 'RARE', 'RBA', 'RBB', 'RBBN', 'RBC', 'RBCAA', 'RBLX', 'RBRK', 'RC', 'RCAT', 'RCEL', 'RCKT', 'RCL', 'RCUS', 'RDDT', 'RDN', 'RDNT', 'RDVT', 'RDW', 'REAL', 'REAX', 'REFI', 'REG', 'REGN', 'RELY', 'REPL', 'REPX', 'RES', 'REX', 'REXR', 'REYN', 'REZI', 'RF', 'RGA', 'RGEN', 'RGLD', 'RGNX', 'RGP', 'RGR', 'RGTI', 'RH', 'RHI', 'RHLD', 'RHP', 'RICK', 'RIG', 'RIGL', 'RIOT', 'RITM', 'RIVN', 'RJF', 'RKLB', 'RKT', 'RL', 'RLAY', 'RLI', 'RLJ', 'RM', 'RMAX', 'RMBS', 'RMD', 'RMNI', 'RMR', 'RNA', 'RNG', 'RNR', 'RNST', 'ROAD', 'ROCK', 'ROG', 'ROIV', 'ROK', 'ROKU', 'ROL', 'ROOT', 'ROP', 'ROST', 'RPAY', 'RPC', 'RPD', 'RPM', 'RPRX', 'RRBI', 'RRC', 'RRR', 'RRX', 'RS', 'RSG', 'RSI', 'RTX', 'RUM', 'RUN', 'RUSHA', 'RUSHB', 'RVLV', 'RVMD', 'RVTY', 'RWT', 'RXO', 'RXRX', 'RXST', 'RYAM', 'RYAN', 'RYN', 'RYTM', 'RYZ', 'RZLV', 'S', 'SABR', 'SAFE', 'SAFT', 'SAH', 'SAIA', 'SAIC', 'SAIL', 'SAM', 'SANA', 'SANM', 'SARO', 'SATS', 'SB', 'SBAC', 'SBCF', 'SBGI', 'SBH', 'SBRA', 'SBSI', 'SBUX', 'SCCO', 'SCHL', 'SCHW', 'SCI', 'SCL', 'SCSC', 'SCVL', 'SD', 'SDGR', 'SDRL', 'SEAT', 'SEB', 'SEE', 'SEG', 'SEI', 'SEIC', 'SEM', 'SEMR', 'SENEA', 'SEPN', 'SERV', 'SEZL', 'SF', 'SFBS', 'SFD', 'SFIX', 'SFL', 'SFM', 'SFNC', 'SFST', 'SG', 'SGHC', 'SGI', 'SGRY', 'SHAK', 'SHBI', 'SHC', 'SHEN', 'SHLS', 'SHO', 'SHOO', 'SHW', 'SIBN', 'SIG', 'SIGA', 'SIGI', 'SILA', 'SION', 'SIRI', 'SITC', 'SITE', 'SITM', 'SJM', 'SKIN', 'SKT', 'SKWD', 'SKY', 'SKYT', 'SKYW', 'SLAB', 'SLB', 'SLDB', 'SLDE', 'SLDP', 'SLG', 'SLGN', 'SLM', 'SLNO', 'SLP', 'SLQT', 'SLS', 'SLVM', 'SM', 'SMA', 'SMBC', 'SMBK', 'SMCI', 'SMG', 'SMMT', 'SMP', 'SMPL', 'SMR', 'SMTC', 'SN', 'SNA', 'SNBR', 'SNCY', 'SNDK', 'SNDR', 'SNDX', 'SNEX', 'SNOW', 'SNPS', 'SNX', 'SO', 'SOC', 'SOFI', 'SOLS', 'SOLV', 'SON', 'SONO', 'SOUN', 'SPB', 'SPFI', 'SPG', 'SPGI', 'SPHR', 'SPNT', 'SPOK', 'SPOT', 'SPRY', 'SPSC', 'SPT', 'SPXC', 'SR', 'SRCE', 'SRE', 'SRPT', 'SRRK', 'SRTA', 'SSB', 'SSD', 'SSNC', 'SSP', 'SSRM', 'SSTI', 'SSTK', 'ST', 'STAA', 'STAG', 'STBA', 'STC', 'STE', 'STEL', 'STEP', 'STGW', 'STKL', 'STLD', 'STNE', 'STNG', 'STOK', 'STRA', 'STRL', 'STRZ', 'STT', 'STWD', 'STZ', 'SUI', 'SUNS', 'SUPN', 'SVC', 'SVRA', 'SVV', 'SW', 'SWBI', 'SWK', 'SWKS', 'SWX', 'SXC', 'SXI', 'SXT', 'SYBT', 'SYF', 'SYK', 'SYNA', 'SYRE', 'SYY', 'T', 'TALK', 'TALO', 'TAP', 'TARS', 'TBBK', 'TBCH', 'TBI', 'TBPH', 'TCBI', 'TCBK', 'TCBX', 'TCMD', 'TCX', 'TDAY', 'TDC', 'TDG', 'TDOC', 'TDS', 'TDUP', 'TDW', 'TDY', 'TE', 'TEAD', 'TEAM', 'TECH', 'TECX', 'TEM', 'TENB', 'TER', 'TERN', 'TEX', 'TFC', 'TFIN', 'TFSL', 'TFX', 'TG', 'TGLS', 'TGT', 'TGTX', 'TH', 'THC', 'THFF', 'THG', 'THO', 'THR', 'THRD', 'THRM', 'THRY', 'TIC', 'TIGO', 'TILE', 'TIPT', 'TITN', 'TJX', 'TK', 'TKO', 'TKR', 'TLN', 'TMCI', 'TMDX', 'TMHC', 'TMO', 'TMP', 'TMUS', 'TNC', 'TNDM', 'TNET', 'TNGX', 'TNK', 'TNL', 'TOL', 'TOST', 'TOWN', 'TPB', 'TPC', 'TPG', 'TPH', 'TPL', 'TPR', 'TR', 'TRC', 'TRDA', 'TREE', 'TREX', 'TRGP', 'TRIP', 'TRMB', 'TRMK', 'TRN', 'TRNO', 'TRNS', 'TROW', 'TROX', 'TRS', 'TRST', 'TRTX', 'TRU', 'TRUP', 'TRV', 'TRVI', 'TSBK', 'TSCO', 'TSHA', 'TSLA', 'TSN', 'TT', 'TTC', 'TTD', 'TTEC', 'TTEK', 'TTGT', 'TTI', 'TTMI', 'TTWO', 'TVTX', 'TW', 'TWI', 'TWLO', 'TWO', 'TWST', 'TXG', 'TXN', 'TXNM', 'TXRH', 'TXT', 'TYL', 'TYRA', 'U', 'UA', 'UAA', 'UAL', 'UAMY', 'UBER', 'UBSI', 'UCB', 'UCTT', 'UDMY', 'UDR', 'UE', 'UEC', 'UFCS', 'UFPI', 'UFPT', 'UGI', 'UHAL', 'UHALB', 'UHS', 'UHT', 'UI', 'UIS', 'ULCC', 'ULH', 'ULTA', 'UMBF', 'UMH', 'UNF', 'UNFI', 'UNH', 'UNIT', 'UNM', 'UNP', 'UNTY', 'UPB', 'UPBD', 'UPS', 'UPST', 'UPWK', 'URBN', 'URGN', 'URI', 'USAR', 'USB', 'USFD', 'USLM', 'USNA', 'USPH', 'UTHR', 'UTI', 'UTL', 'UTMD', 'UTZ', 'UUUU', 'UVE', 'UVSP', 'UVV', 'UWMC', 'V', 'VAC', 'VAL', 'VC', 'VCEL', 'VCTR', 'VCYT', 'VECO', 'VEEV', 'VEL', 'VERA', 'VERX', 'VFC', 'VIAV', 'VICI', 'VICR', 'VIK', 'VIR', 'VIRT', 'VISN', 'VITL', 'VKTX', 'VLO', 'VLTO', 'VLY', 'VMC', 'VMD', 'VMI', 'VNDA', 'VNO', 'VNOM', 'VNT', 'VOYA', 'VOYG', 'VPG', 'VRDN', 'VRE', 'VREX', 'VRNS', 'VRRM', 'VRSK', 'VRSN', 'VRT', 'VRTS', 'VRTX', 'VSAT', 'VSCO', 'VSEC', 'VSH', 'VSNT', 'VST', 'VSTM', 'VSTS', 'VTOL', 'VTR', 'VTRS', 'VTS', 'VVV', 'VVX', 'VYGR', 'VYX', 'VZ', 'W', 'WAB', 'WABC', 'WAFD', 'WAL', 'WALD', 'WASH', 'WAT', 'WAY', 'WBD', 'WBS', 'WCC', 'WD', 'WDAY', 'WDC', 'WDFC', 'WEAV', 'WEC', 'WELL', 'WEN', 'WERN', 'WEST', 'WEX', 'WFC', 'WFRD', 'WGO', 'WGS', 'WH', 'WHD', 'WHR', 'WINA', 'WING', 'WK', 'WKC', 'WLDN', 'WLFC', 'WLK', 'WLY', 'WM', 'WMB', 'WMK', 'WMS', 'WMT', 'WNC', 'WOOF', 'WOR', 'WPC', 'WRB', 'WRBY', 'WRLD', 'WS', 'WSBC', 'WSBF', 'WSC', 'WSFS', 'WSM', 'WSO', 'WSR', 'WST', 'WT', 'WTBA', 'WTFC', 'WTI', 'WTM', 'WTRG', 'WTS', 'WTTR', 'WTW', 'WU', 'WULF', 'WVE', 'WWD', 'WWW', 'WY', 'WYNN', 'XEL', 'XENE', 'XERS', 'XHR', 'XMTR', 'XNCR', 'XOM', 'XP', 'XPEL', 'XPER', 'XPO', 'XPOF', 'XPRO', 'XRAY', 'XRN', 'XRX', 'XYL', 'XYZ', 'YELP', 'YETI', 'YEXT', 'YORW', 'YOU', 'YUM', 'Z', 'ZBH', 'ZBIO', 'ZBRA', 'ZD', 'ZETA', 'ZG', 'ZION', 'ZIP', 'ZM', 'ZS', 'ZTS', 'ZUMZ', 'ZVRA', 'ZWS', 'ZYME']

# Assurance-Vie Compliant : 
# TICKERS = ["AMUN.PA","RF.PA","LDO.MI","NEX.PA","PLTR","RHM.DE","RR.L","GLE.PA","BBVA.MC","BAYN.DE","UCG.MI","MT","HO.PA","SPIE.PA","INGA.AS","EN.PA","ISP.MI","NDA-FI.HE","ORA.PA","ENGI.PA","GOOGL","FGR.PA","SAF.PA","CAT","IBE.MC","AM.PA","PRX.AS","DPW.DE","RXL.PA","BNP.PA","ASML.AS","GS","LR.PA","ACA.PA","AVGO","ALV.DE","NOK","ENEL.MI","AIR.PA","URW.PA","SIE.DE","ERF.PA","JNJ","KER.PA","AZN","FR.PA","ENI.MI","NVDA","GTT.PA","SCR.PA","LI.PA","DG.PA","IFX.DE","CS.PA","IBM","JPM","ENX.PA","BMW.DE","BN.PA","ALO.PA","VOW3.DE","MUV2.DE","CSCO","EL.PA","ABI.BR","ITX.MC","MBG.DE","DIM.PA","AMGN","AD.AS","AXP","MMM","VIE.PA","WMT","BA","NFLX","OR.PA","BIM.PA","TRV","BAS.DE","TTE","CA.PA","AC.PA","AI.PA","GET.PA","SGO.PA","MSFT","DB1.DE","ADP.PA","META","KO","TSLA","V","SU.PA","AAPL","ISRG","DTE.DE","ADYEN.AS","MRK","MCD","CVX","AMZN","BVI.PA","STM","RMS.PA","DIS","LIN","VZ","CAP.PA","GFC.PA","ML.PA","SAN.PA","SAP.DE","PUB.PA","PEP","COST","TMUS","HD","RACE","HON","PG","RNO.PA","NKE","TEP.PA","STLA","ADS.DE","DSY.PA","AKE.PA","CRM","RI.PA","EDEN.PA","UNH","WKL.AS","SW.PA""MU", "SNDK", "TEP.PA", "ALSTI.PA", "NVDA", "GOOG", "AMZN", "META", "BNP.PA", "CRM", "ADBE", "AMD", "QCOM"]

TICKERS = ["MU", "SNDK", "TEP.PA", "ALSTI.PA", "NVDA", "GOOG", "AMZN", "META", "BNP.PA", "CRM", "ADBE", "AMD", "QCOM"]
# TICKERS = ["BA","LMT","RTX","NOC","GD","LHX","AVAV","AIR","BAESY","CAE","TXT","KTOS","ONDS","RCAT","PARRO","EH","FLT","ESLT","NXSN"]


BASE_DIR = Path(__file__).resolve().parent
FILENAME = BASE_DIR / "Holden_Model_MU.xlsx"

USER_AGENT = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'}

# --- MODEL TUNING CONSTANTS ---
# Insider Conviction Scoring Thresholds, based on the existing academic literature.
INSIDER_DOLLAR_LARGE = 1_000_000
INSIDER_PCT_LARGE = 0.0002
INSIDER_DOLLAR_MODERATE = 250_000
INSIDER_PCT_MODERATE = 0.00005

INSIDER_STAKE_PCT_FOR_SCORE_4 = 100
INSIDER_STAKE_PCT_FOR_SCORE_3 = 50
INSIDER_STAKE_PCT_FOR_SCORE_2 = 20
INSIDER_STAKE_PCT_FOR_SCORE_1 = 5

MIN_HISTORICAL_PE_LOW = 5.0
MIN_HISTORICAL_PE_HIGH = 10.0

ROLLING_PE_PERIOD = "2y"
GAAP_REPORTING_DELAY_DAYS = 45
MAX_VALID_PE = 300
PE_LOW_QUANTILE = 0.05
PE_HIGH_QUANTILE = 0.95
PE_LOW_HIST_FALLBACK_MULT = 0.9
CAGR_YEARS = 3
OPENINSIDER_DAYS = 180

CURRENCY_MAP = {
    'USD': '$', 'EUR': '€', 'GBP': '£',
    'JPY': '¥', 'CNY': '¥', 'INR': '₹',
    'CAD': 'C$', 'AUD': 'A$'
}

REGION_MAP = {
    'United States': 'North America', 'Canada': 'North America',
    'France': 'Europe', 'Germany': 'Europe', 'United Kingdom': 'Europe',
    'Spain': 'Europe', 'Italy': 'Europe', 'Netherlands': 'Europe',
    'Japan': 'Asia-Pac', 'China': 'Asia-Pac', 'India': 'Asia-Pac',
    'Taiwan': 'Asia-Pac', 'Australia': 'Asia-Pac'
}


# --- REGIME CLASSIFICATION ---
def classify_regime(eps_trend, pe_trend, fwd_growth):
    """Classify stock into ERG+ regime based on directional signals."""
    THRESHOLD = 0.03
    e_up = eps_trend > THRESHOLD
    p_up = pe_trend > THRESHOLD
    p_down = pe_trend < -THRESHOLD
    f_up = fwd_growth > THRESHOLD
    f_down = fwd_growth < -THRESHOLD

    if e_up:
        if p_down:
            if f_up:     return "Golden Gap", "\U0001f7e2 Strong Opportunity"
            elif f_down: return "Value Trap (Peak)", "\U0001f534 Avoid"
            else:        return "Value Trap Risk", "\U0001f7e1 Investigate"
        elif p_up:
            if f_up:     return "Growth Expansion", "\U0001f7e2 Momentum"
            elif f_down: return "Late-Cycle Excess", "\U0001f7e1 Trim"
            else:        return "Late-Cycle Excess", "\U0001f7e1 Trim"
        else:
            if f_up:     return "Confirmed Growth", "\U0001f7e2 Momentum"
            elif f_down: return "Decelerating", "\U0001f7e1 Investigate"
            else:        return "Growth Stalling", "\U0001f7e1 Investigate"
    else:
        if p_down:
            if f_up:     return "Turnaround", "\U0001f7e1 Speculative"
            elif f_down: return "Decline", "\U0001f534 Avoid"
            else:        return "Decline", "\U0001f534 Avoid"
        elif p_up:
            if f_up:     return "Recovery Expected", "\U0001f7e1 Early Entry"
            elif f_down: return "Overvalued", "\U0001f534 Avoid"
            else:        return "Stagnation", "\U0001f7e1 Investigate"
        else:
            if f_up:     return "Turnaround", "\U0001f7e1 Speculative"
            elif f_down: return "Decline", "\U0001f534 Avoid"
            else:        return "Stagnation", "\u26aa Neutral"


# --- DATA GATHERING ---
def get_insider_data(ticker):
    """
    Scrapes OpenInsider for the last 6 months.
    Returns dictionary with Net Buying, Unique Buyers count, and Average Stake Increase %.
    """
    clean_sym = ticker.split('.')[0]
    print(f"  [{clean_sym}] Fetching OpenInsider data (L6M)...")

    url = f"http://openinsider.com/screener?s={clean_sym}&o=&pl=&ph=&ll=&lh=&fd={OPENINSIDER_DAYS}&fdr=&td=&tdr=&fdlyl=&fdlyh=&daysago=&xp=1&xs=1&vl=&vh=&ocl=&och=&sic1=-1&sicl=100&sich=9999&grp=0&nfl=&nfh=&nil=&nih=&nol=&noh=&v2l=&v2h=&oc2l=&oc2h=&sortcol=0&cnt=100&page=1"

    default_res = {'net_buying': 0.0, 'unique_buyers': 0, 'avg_stake_inc': 0.0}

    try:
        response = requests.get(url, headers=USER_AGENT, timeout=10)
        if response.status_code != 200:
            print(f"  [{clean_sym}] OpenInsider request failed with status code: {response.status_code}")
            return default_res

        soup = BeautifulSoup(response.text, 'html.parser')
        table = soup.find('table', {'class': 'tinytable'})

        if not table:
            print(f"  [{clean_sym}] No insider table found. Assuming 0 activity or foreign ticker.")
            return default_res

        net_buying = 0.0
        unique_buyers_set = set()
        stake_increases = []

        rows = table.find('tbody').find_all('tr')
        valid_trades = 0

        for row in rows:
            cols = row.find_all('td')
            if len(cols) < 12: continue

            insider_name = cols[4].text.strip()
            trade_type = cols[6].text.strip()
            delta_own_txt = cols[10].text.strip().replace('%', '').replace('+', '').replace(',', '')
            value_txt = cols[11].text.strip().replace('$', '').replace(',', '').replace('+', '').replace('-', '')

            try:
                val = abs(float(value_txt))
                valid_trades += 1
            except ValueError:
                val = 0.0

            if 'Purchase' in trade_type:
                net_buying += val
                unique_buyers_set.add(insider_name)

                # Parse Stake Increase (capped at 100% to prevent extreme skewing from "New" positions)
                if delta_own_txt.lower() == 'new' or '>999' in delta_own_txt:
                    stake_inc = 100.0
                else:
                    try:
                        stake_inc = float(delta_own_txt)
                        stake_inc = min(stake_inc, 100.0)  # Cap at 100%
                    except ValueError:
                        stake_inc = 0.0

                if stake_inc > 0:
                    stake_increases.append(stake_inc)

            elif 'Sale' in trade_type:
                net_buying -= val

        unique_buyers_count = len(unique_buyers_set)
        avg_stake_inc = sum(stake_increases) / len(stake_increases) if stake_increases else 0.0

        print(
            f"  [{clean_sym}] Parsed {valid_trades} trades. Net Buy: ${net_buying:,.0f} | Unique Buyers: {unique_buyers_count} | Avg Stake Inc: {avg_stake_inc:.1f}%")
        return {'net_buying': net_buying, 'unique_buyers': unique_buyers_count, 'avg_stake_inc': avg_stake_inc}

    except Exception as e:
        print(f"  [{clean_sym}] OpenInsider scraping error: {e}")
        return default_res


def get_data(ticker):
    """Fetches Basic (GAAP) and reconstructs Street (Adjusted) EPS from earnings history, plus Next FY estimates."""
    clean_ticker = ticker.split()[0]
    print(f"\n[{clean_ticker}] Starting data fetch process...")

    try:
        stock = yf.Ticker(clean_ticker)
        info = stock.info
        print(f"  [{clean_ticker}] Successfully retrieved yfinance info dictionary.")
    except Exception as e:
        print(f"  [{clean_ticker}] CRITICAL ERROR: Failed to instantiate yfinance Ticker or retrieve info. ({e})")
        return None

    company_name = info.get('longName', info.get('shortName', clean_ticker))
    market_cap = info.get('marketCap', 0.0)
    currency_code = info.get('currency', 'USD').upper()
    curr_price = info.get('currentPrice', info.get('regularMarketPrice', 0.0))
    currency_symbol = CURRENCY_MAP.get(currency_code, '$')
    sector = info.get('sector', 'Unknown')
    country = info.get('country', 'Unknown')
    region = REGION_MAP.get(country, country)

    if curr_price == 0: return None

    eps_basic_ttm = info.get('trailingEps')
    if eps_basic_ttm is None: eps_basic_ttm = info.get('dilutedEpsTrailingTwelveMonths', 0.0)

    eps_basic_prior = eps_basic_ttm
    try:
        fin = stock.financials
        if 'Basic EPS' in fin.index and len(fin.columns) >= 2:
            val = fin.loc['Basic EPS'].iloc[1]
            if not np.isnan(val): eps_basic_prior = val
    except:
        pass

    eps_street_ttm = eps_basic_ttm
    eps_street_prior = eps_basic_prior
    surprise_display_val = 0.0

    try:
        dates = stock.earnings_dates
        if dates is not None and not dates.empty:
            dates.index = dates.index.tz_localize(None)
            actual_col = next((col for col in dates.columns if 'Actual' in str(col) or 'Reported' in str(col)), None)

            if actual_col:
                now = pd.Timestamp.now()
                past = dates[dates.index < now].dropna(subset=[actual_col]).sort_index(ascending=False)

                if len(past) >= 4:
                    last_4 = past.head(4)
                    if last_4[actual_col].sum() != 0: eps_street_ttm = last_4[actual_col].sum()
                    if 'Surprise(%)' in dates.columns:
                        surp = last_4['Surprise(%)'].dropna()
                        surprise_display_val = surp.abs().mean() if (surp > 0).any() and (
                                    surp < 0).any() else surp.mean()

                if len(past) >= 8:
                    prior_4 = past.iloc[4:8]
                    if prior_4[actual_col].sum() != 0: eps_street_prior = prior_4[actual_col].sum()
    except:
        pass

    analyst_count = info.get('numberOfAnalystOpinions', 0)
    est_year_str = "Current FY"
    est_year_str_nxt = "Next FY"

    try:
        nxt_fy_ts = info.get('nextFiscalYearEnd')
        if nxt_fy_ts:
            est_year_dt = pd.to_datetime(nxt_fy_ts, unit='s')
            est_year_str = f"FY {est_year_dt.year} ({est_year_dt.strftime('%b')})"
            nxt_year_dt = est_year_dt + pd.DateOffset(years=1)
            est_year_str_nxt = f"FY {nxt_year_dt.year} ({nxt_year_dt.strftime('%b')})"
    except:
        pass

    eps_mid, eps_high, eps_low = 0, 0, 0
    eps_mid_nxt, eps_high_nxt, eps_low_nxt = 0, 0, 0
    base_eps = eps_street_ttm if eps_street_ttm > 0 else eps_basic_ttm

    try:
        est = stock.earnings_estimate
        if est is not None and '0y' in est.index:
            eps_mid = est.loc['0y', 'avg']
            eps_high = est.loc['0y', 'high']
            eps_low = est.loc['0y', 'low']
            if abs(eps_high - eps_low) < 0.001:
                eps_low, eps_high = eps_mid * 0.75, eps_mid * 1.25
        else:
            raise ValueError

        if est is not None and '1y' in est.index:
            eps_mid_nxt = est.loc['1y', 'avg']
            eps_high_nxt = est.loc['1y', 'high']
            eps_low_nxt = est.loc['1y', 'low']
            if abs(eps_high_nxt - eps_low_nxt) < 0.001:
                eps_low_nxt, eps_high_nxt = eps_mid_nxt * 0.75, eps_mid_nxt * 1.25
        else:
            eps_mid_nxt = eps_mid * 1.10
            eps_high_nxt, eps_low_nxt = eps_mid_nxt * 1.25, eps_mid_nxt * 0.90

    except:
        eps_mid = base_eps * 1.10
        eps_high, eps_low = eps_mid * 1.25, eps_mid * 0.90
        eps_mid_nxt = eps_mid * 1.10
        eps_high_nxt, eps_low_nxt = eps_mid_nxt * 1.25, eps_mid_nxt * 0.90

    pe_current = curr_price / eps_basic_ttm if (eps_basic_ttm and eps_basic_ttm != 0) else 0
    pe_low_hist = max(MIN_HISTORICAL_PE_LOW, pe_current * 0.7)
    pe_high_hist = max(MIN_HISTORICAL_PE_HIGH, pe_current * 1.3)

    # --- Historical 1Y Trend variables (computed inside rolling PE block) ---
    perf_1y = 0.0
    eps_trend_1y = 0.0
    pe_trend_1y = 0.0

    # Calculate ROLLING Historical P/E Ranges safely via yfinance
    try:
        hist = stock.history(period=ROLLING_PE_PERIOD)
        if not hist.empty:
            hist.index = hist.index.tz_localize(None)
            closes = hist['Close']

            # 1) Rolling GAAP
            ttm_gaap_series = pd.Series(index=closes.index, dtype=float)
            q_fin = stock.quarterly_financials
            if q_fin is not None and 'Basic EPS' in q_fin.index:
                basic_eps = q_fin.loc['Basic EPS'].dropna().sort_index(ascending=True)
                rolling_4q_gaap = basic_eps.rolling(4).sum().dropna()
                if not rolling_4q_gaap.empty:
                    # Delay by GAAP_REPORTING_DELAY_DAYS to avoid look-ahead bias
                    gaap_avail_dates = rolling_4q_gaap.index + pd.Timedelta(days=GAAP_REPORTING_DELAY_DAYS)
                    gaap_ttm_df = pd.DataFrame({'eps': rolling_4q_gaap.values}, index=gaap_avail_dates).sort_index()
                    gaap_ttm_df = gaap_ttm_df[~gaap_ttm_df.index.duplicated(keep='last')]
                    ttm_gaap_series = gaap_ttm_df['eps'].reindex(closes.index, method='ffill')

            # 2) Rolling Street
            ttm_street_series = pd.Series(index=closes.index, dtype=float)
            dates = stock.earnings_dates
            if dates is not None and not dates.empty:
                dates = dates.copy()
                dates.index = dates.index.tz_localize(None)
                actual_col = next((col for col in dates.columns if 'Actual' in str(col) or 'Reported' in str(col)), None)
                if actual_col:
                    actuals = dates[actual_col].dropna().sort_index(ascending=True)
                    rolling_4q_street = actuals.rolling(4).sum().dropna()
                    if not rolling_4q_street.empty:
                        street_ttm_df = pd.DataFrame({'eps': rolling_4q_street.values}, index=rolling_4q_street.index).sort_index()
                        street_ttm_df = street_ttm_df[~street_ttm_df.index.duplicated(keep='last')]
                        ttm_street_series = street_ttm_df['eps'].reindex(closes.index, method='ffill')

            daily_street_pe = closes / ttm_street_series
            daily_gaap_pe = closes / ttm_gaap_series

            daily_pe = daily_street_pe if not daily_street_pe.dropna().empty else daily_gaap_pe
            daily_pe = daily_pe[(daily_pe > 0) & (daily_pe < MAX_VALID_PE)].dropna()

            if not daily_pe.empty:
                pe_low_hist = daily_pe.quantile(PE_LOW_QUANTILE)
                pe_high_hist = daily_pe.quantile(PE_HIGH_QUANTILE)

            # --- Compute Historical 1Y Trends from the rolling series ---
            one_year_ago = closes.index[-1] - pd.DateOffset(years=1)

            # 1Y Price Performance
            closes_1y = closes[closes.index >= one_year_ago]
            if len(closes_1y) >= 2:
                perf_1y = (closes_1y.iloc[-1] / closes_1y.iloc[0]) - 1

            # 1Y EPS Trend (prefer street, fall back to GAAP)
            eps_series = ttm_street_series if not ttm_street_series.dropna().empty else ttm_gaap_series
            eps_1y = eps_series[eps_series.index >= one_year_ago].dropna()
            if len(eps_1y) >= 2 and eps_1y.iloc[0] != 0:
                eps_trend_1y = (eps_1y.iloc[-1] / eps_1y.iloc[0]) - 1
            elif eps_street_prior and eps_street_prior != 0:
                eps_trend_1y = (eps_street_ttm / eps_street_prior) - 1

            # 1Y P/E Trend (from daily PE series, or derive via identity)
            pe_1y = daily_pe[daily_pe.index >= one_year_ago]
            if len(pe_1y) >= 2 and pe_1y.iloc[0] != 0:
                pe_trend_1y = (pe_1y.iloc[-1] / pe_1y.iloc[0]) - 1
            elif (1 + eps_trend_1y) != 0:
                pe_trend_1y = ((1 + perf_1y) / (1 + eps_trend_1y)) - 1

    except Exception as e:
        print(f"  [{clean_ticker}] Rolling PE calculation fallback failed: {e}")

    # Only allow fallback when PE is positive and genuinely below historical range
    if 0 < pe_current < pe_low_hist:
        pe_low_hist = max(MIN_HISTORICAL_PE_LOW, pe_current * PE_LOW_HIST_FALLBACK_MULT)

    # Final safety floor: PE bounds must never fall below configured minimums
    pe_low_hist = max(MIN_HISTORICAL_PE_LOW, pe_low_hist)
    pe_high_hist = max(MIN_HISTORICAL_PE_HIGH, pe_high_hist)

    cagr_3y = 0.0
    try:
        fin_annual = stock.financials
        if 'Basic EPS' in fin_annual.index:
            eps_years = fin_annual.loc['Basic EPS'].sort_index(ascending=True)
            if len(eps_years) >= (CAGR_YEARS + 1) and eps_years.iloc[-(CAGR_YEARS + 1)] > 0 and eps_years.iloc[-1] > 0:
                cagr_3y = ((eps_years.iloc[-1] / eps_years.iloc[-(CAGR_YEARS + 1)]) ** (1 / CAGR_YEARS)) - 1
    except:
        pass

    # --- Forward-Confirmed ERG+ ---
    base_eps_fwd = eps_street_ttm if eps_street_ttm > 0 else eps_basic_ttm
    if base_eps_fwd and base_eps_fwd > 0 and eps_mid and eps_mid > 0:
        implied_fwd_growth = (eps_mid - base_eps_fwd) / base_eps_fwd
    else:
        implied_fwd_growth = 0.0

    erg_raw = (eps_trend_1y - perf_1y) if eps_trend_1y > 0 else 0.0
    if eps_trend_1y > 0.01:
        fcr_raw = implied_fwd_growth / eps_trend_1y
    else:
        fcr_raw = 0.0
    fcr = max(-1.0, min(2.0, fcr_raw))
    erg_plus = erg_raw * max(fcr, 0)

    regime, regime_signal = classify_regime(eps_trend_1y, pe_trend_1y, implied_fwd_growth)

    insider_data = get_insider_data(clean_ticker)

    # --- Quality & Risk Metrics ---
    profit_margin = info.get('profitMargins', 0.0)
    if profit_margin is None: profit_margin = 0.0
    
    fcf = info.get('freeCashflow', 0.0)
    if fcf is None: fcf = 0.0
    fcf_yield = (fcf / market_cap) if market_cap and market_cap > 0 else 0.0
    
    raw_de = info.get('debtToEquity', None)
    if raw_de is not None:
        de_ratio = raw_de / 100.0  # yfinance returns percentage e.g., 85.5 for 0.85x
    else:
        de_ratio = -1.0 # Use -1 to denote missing
        
    roe = info.get('returnOnEquity', 0.0)
    if roe is None: roe = 0.0

    return {
        'ticker': clean_ticker, 'company_name': company_name, 'market_cap': market_cap,
        'currency': currency_symbol, 'price': curr_price,
        'sector': sector, 'region': region,
        'est_year_str': est_year_str, 'est_year_str_nxt': est_year_str_nxt,
        'eps_basic_ttm': eps_basic_ttm, 'eps_basic_prior': eps_basic_prior,
        'eps_street_ttm': eps_street_ttm, 'eps_street_prior': eps_street_prior,
        'eps_low': eps_low, 'eps_mid': eps_mid, 'eps_high': eps_high,
        'eps_low_nxt': eps_low_nxt, 'eps_mid_nxt': eps_mid_nxt, 'eps_high_nxt': eps_high_nxt,
        'pe_current': pe_current, 'pe_low_hist': pe_low_hist, 'pe_high_hist': pe_high_hist,
        'cagr_3y': cagr_3y, 'surprise_avg': surprise_display_val, 'analysts': analyst_count,
        'insider_net': insider_data['net_buying'], 'insider_buy_count': insider_data['unique_buyers'],
        'insider_avg_stake_inc': insider_data['avg_stake_inc'],
        'perf_1y': perf_1y, 'eps_trend_1y': eps_trend_1y, 'pe_trend_1y': pe_trend_1y,
        'implied_fwd_growth': implied_fwd_growth, 'fcr': fcr, 'erg_plus': erg_plus,
        'regime': regime, 'regime_signal': regime_signal,
        'profit_margin': profit_margin, 'fcf_yield': fcf_yield, 'de_ratio': de_ratio, 'roe': roe
    }

# Copyright (c) 2026 Dylan H Wilding. All rights reserved.
#
# This source code and the Holden Valuation Model are proprietary and confidential.
# Unauthorized copying, distribution, modification, reverse engineering, or use of
# this file, in whole or in part, is prohibited without prior written permission
# from Dylan H Wilding.
#
# This model is provided for informational and research purposes only and does not
# constitute investment, legal, tax, or accounting advice. No warranty is provided
# as to the accuracy, completeness, or fitness for any particular purpose of the
# data, calculations, forecasts, or outputs generated by this model. Use at your own risk.

# --- FORMATTING ---
def get_formats(workbook, symbol):
    base = {'font_name': 'Arial', 'font_size': 10, 'align': 'center', 'valign': 'vcenter'}
    num_fmt = f'"{symbol}"#,##0.00'

    def add(props): return workbook.add_format({**base, **props})

    return {
        'title': add({'bold': True, 'bg_color': '#006100', 'font_color': 'white', 'border': 1, 'font_size': 12}),
        'pe_label': add({'bold': True, 'bg_color': '#92D050', 'border': 1}),
        'pe_val_base': add({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'num_format': '0.0x'}),
        'pe_val_mid': add({'bold': True, 'bg_color': '#A6A6A6', 'border': 1, 'num_format': '0.0x'}),
        'eps_label': add({'bold': True, 'bg_color': '#92D050', 'border': 1, 'rotation': 90}),
        'eps_val': add({'bold': True, 'bg_color': 'white', 'border': 1, 'num_format': num_fmt}),
        'eps_mid': add({'bold': True, 'bg_color': '#FFCCFF', 'border': 1, 'num_format': num_fmt}),
        'eps_low': add({'bold': True, 'bg_color': '#CCFFFF', 'border': 1, 'num_format': num_fmt}),
        'outer_zone': add({'bg_color': '#FCE4D6', 'border': 1, 'num_format': '0'}),
        'mid_zone': add({'bg_color': '#FFFFCC', 'border': 1, 'num_format': '0'}),
        'center_zone': add({'bg_color': '#E2EFDA', 'border': 1, 'num_format': '0'}),
        'outer_zone_pct': add({'bg_color': '#FCE4D6', 'border': 1, 'num_format': '0%', 'align': 'center'}),
        'mid_zone_pct': add({'bg_color': '#FFFFCC', 'border': 1, 'num_format': '0%', 'align': 'center'}),
        'center_zone_pct': add({'bg_color': '#E2EFDA', 'border': 1, 'num_format': '0%', 'align': 'center'}),
        'holden_outer': add({'bg_color': '#FCE4D6', 'border': 1, 'num_format': '0.0%'}),
        'holden_mid': add({'bg_color': '#FFFFCC', 'border': 1, 'num_format': '0.0%'}),
        'holden_center': add({'bg_color': '#E2EFDA', 'border': 1, 'num_format': '0.0%'}),
        'stat_head': add({'bold': True, 'align': 'left', 'bottom': 1, 'bg_color': '#E7E6E6'}),
        'stat_subhead': add({'bold': True, 'align': 'left', 'italic': True, 'font_color': '#595959', 'bottom': 1}),
        'stat_label': add({'align': 'left'}),
        'stat_val': add({'bold': True, 'num_format': num_fmt, 'align': 'right'}),
        'stat_val_txt': add({'bold': True, 'align': 'right'}),
        'stat_val_mcap': add({'bold': True, 'num_format': '"$"#,##0', 'align': 'right'}),
        'stat_val_int': add({'bold': True, 'num_format': '#,##0', 'align': 'right'}),
        'stat_val_score_int': add(
            {'bold': True, 'num_format': '0', 'align': 'right', 'bg_color': '#D9E1F2', 'border': 1}),
        'stat_val_pe': add({'bold': True, 'num_format': '0.00x', 'align': 'right'}),
        'stat_val_score': add({'bold': True, 'num_format': '0.0%', 'align': 'right'}),
        'stat_val_fcr': add({'bold': True, 'num_format': '0.00', 'align': 'right'}),
        'stat_val_real': add(
            {'bold': True, 'num_format': '0.0%', 'align': 'right', 'bg_color': '#E2EFDA', 'border': 1}),
        'stat_val_blue': add(
            {'bold': True, 'num_format': num_fmt, 'align': 'right', 'bg_color': '#CCFFFF', 'border': 1}),
        'stat_val_purple': add(
            {'bold': True, 'num_format': num_fmt, 'align': 'right', 'bg_color': '#FFCCFF', 'border': 1}),
        'stat_val_pe_grey': add(
            {'bold': True, 'num_format': '0.00x', 'align': 'right', 'bg_color': '#A6A6A6', 'border': 1}),
        'est_growth_base': add({'italic': True, 'font_color': '#006100', 'num_format': '+0.0%', 'align': 'left'}),
        'est_growth_neg': add({'italic': True, 'font_color': '#9C0006', 'num_format': '0.0%', 'align': 'left'}),
        'diag_ok': add({'bold': True, 'font_color': '#006100', 'align': 'right'}),
        'diag_warn': add({'bold': True, 'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'align': 'right', 'border': 1}),
        'diag_mid': add({'bold': True, 'bg_color': '#FFEB9C', 'font_color': '#9C6500', 'align': 'right', 'border': 1}),

        # Insider Specific Formatting
        'insider_pos': add(
            {'bold': True, 'font_color': '#006100', 'bg_color': '#C6EFCE', 'num_format': '"$"#,##0', 'align': 'right',
             'border': 1}),
        'insider_neg': add(
            {'bold': True, 'font_color': '#9C0006', 'bg_color': '#FFC7CE', 'num_format': '"$"#,##0', 'align': 'right',
             'border': 1}),
        'insider_neutral': add(
            {'bold': True, 'font_color': '#595959', 'bg_color': '#F2F2F2', 'num_format': '"$"#,##0', 'align': 'right',
             'border': 1}),

        'peg_deep_green': add({'bg_color': '#006100', 'font_color': 'white', 'border': 1, 'num_format': '0.00'}),
        'peg_green': add({'bg_color': '#C6EFCE', 'font_color': 'black', 'border': 1, 'num_format': '0.00'}),
        'peg_yellow': add({'bg_color': '#FFEB9C', 'font_color': 'black', 'border': 1, 'num_format': '0.00'}),
        'peg_orange': add({'bg_color': '#FFCC99', 'font_color': 'black', 'border': 1, 'num_format': '0.00'}),
        'peg_red': add({'bg_color': '#FFC7CE', 'font_color': 'black', 'border': 1, 'num_format': '0.00'}),
        'peg_black': add({'bg_color': '#000000', 'font_color': 'white', 'border': 1, 'num_format': '0.00'}),
        'peg_nm': add({'bg_color': '#F2F2F2', 'font_color': 'gray', 'border': 1}),

        # --- Quality & Risk Tiers ---
        'tier_dkgreen': add({'bold': True, 'bg_color': '#006100', 'font_color': 'white', 'num_format': '0.0%', 'align': 'right', 'border': 1}),
        'tier_green': add({'bold': True, 'bg_color': '#C6EFCE', 'font_color': '#006100', 'num_format': '0.0%', 'align': 'right', 'border': 1}),
        'tier_yellow': add({'bold': True, 'bg_color': '#FFEB9C', 'font_color': '#9C6500', 'num_format': '0.0%', 'align': 'right', 'border': 1}),
        'tier_red': add({'bold': True, 'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'num_format': '0.0%', 'align': 'right', 'border': 1}),
        'tier_black': add({'bold': True, 'bg_color': '#000000', 'font_color': 'white', 'num_format': '0.0%', 'align': 'right', 'border': 1}),
        
        'tier_de_dkgreen': add({'bold': True, 'bg_color': '#006100', 'font_color': 'white', 'num_format': '0.00x', 'align': 'right', 'border': 1}),
        'tier_de_green': add({'bold': True, 'bg_color': '#C6EFCE', 'font_color': '#006100', 'num_format': '0.00x', 'align': 'right', 'border': 1}),
        'tier_de_yellow': add({'bold': True, 'bg_color': '#FFEB9C', 'font_color': '#9C6500', 'num_format': '0.00x', 'align': 'right', 'border': 1}),
        'tier_de_red': add({'bold': True, 'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'num_format': '0.00x', 'align': 'right', 'border': 1}),
        'tier_de_black': add({'bold': True, 'bg_color': '#000000', 'font_color': 'white', 'num_format': '0.00x', 'align': 'right', 'border': 1}),

        'resilience_good': add(
            {'bold': True, 'bg_color': '#C6EFCE', 'font_color': '#006100', 'num_format': '0.00', 'align': 'right',
             'border': 1}),
        'resilience_bad': add(
            {'bold': True, 'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'num_format': '0.00', 'align': 'right',
             'border': 1}),
        'cushion_val': add({'font_color': '#006100', 'num_format': '0.0%', 'align': 'right'}),
        'risk_val': add({'font_color': '#9C0006', 'num_format': '0.0%', 'align': 'right'}),
        'legend_bold': add({'bold': True, 'font_size': 9, 'align': 'left'}),
        'legend_norm': add({'font_size': 9, 'align': 'left', 'italic': True, 'font_color': '#595959'}),

        'input_header': add({'bold': True, 'bg_color': '#808080', 'font_color': 'white', 'border': 1}),
        'dash_label': add({'bold': True, 'bg_color': '#E7E6E6', 'align': 'left', 'border': 1}),
        'dash_input': add({'bg_color': '#FFF2CC', 'border': 1, 'align': 'left'}),

        'comp_head': add({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1}),
        'comp_ticker': add({'bold': True, 'align': 'left'}),
        'comp_txt': add({'align': 'left'}),
        'comp_mcap': add({'num_format': '"$"#,##0'}),
        'comp_num': add({'num_format': '#,##0.00'}),
        'comp_pct': add({'num_format': '0.0%'}),
        'comp_pe': add({'num_format': '0.0x'}),
        'comp_dollar': add({'num_format': '"$"#,##0'}),
        'comp_int': add({'num_format': '#,##0'})
    }


# --- MAIN LOGIC ---
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Holden Valuation Model")
    parser.add_argument("--tickers", nargs="+", help="List of tickers to process (e.g. ABC XYZ)")
    parser.add_argument("--output", type=str, help="Output Excel filename (e.g. custom.xlsx)")
    args = parser.parse_args()

    if args.tickers:
        TICKERS = args.tickers
    if args.output:
        FILENAME = BASE_DIR / args.output

    # Deduplicate TICKERS while preserving order
    TICKERS = list(dict.fromkeys(TICKERS))

    ALL_DATA = []

    print("\n==============================")
    print("Initiating Batch Data Sequence")
    print("==============================\n")

    for t in TICKERS:
        d = get_data(t)
        if d:
            ALL_DATA.append(d)
        else:
            print(f"[{t}] FAILED. Moving to next ticker.\n")

    print("\nGenerating Dashboards...")
    writer = pd.ExcelWriter(
        FILENAME,
        engine='xlsxwriter',
        engine_kwargs={'options': {'nan_inf_to_errors': True}}
    )
    workbook = writer.book

    # 1. GENERATE DASHBOARDS
    data_start_row = 10
    used_sheet_names = set()

    for i, d in enumerate(ALL_DATA):
        ticker = d['ticker']
        print(f"  Formatting dashboard for {ticker}...")

        # --- Unique and Valid Sheet Naming ---
        # Excel forbidden: \ / ? * : [ ]
        clean_name = ticker.replace('\\','').replace('/','').replace('?','').replace('*','').replace(':','').replace('[','').replace(']','')
        base_name = clean_name[:31] # Excel limit
        
        sheet_name = base_name
        counter = 1
        while sheet_name.lower() in used_sheet_names:
            suffix = f"_{counter}"
            sheet_name = f"{base_name[:31-len(suffix)]}{suffix}"
            counter += 1
        
        used_sheet_names.add(sheet_name.lower())
        
        sheet = workbook.add_worksheet(sheet_name)
        sheet.hide_gridlines(2)
        fmt = get_formats(workbook, d['currency'])

        inp_row = data_start_row + 1 + i


        def get_ref(col_idx):
            return f"Inputs!{xl_rowcol_to_cell(inp_row, col_idx, row_abs=True, col_abs=True)}"


        # Mapping References
        ref_ticker = get_ref(0)
        ref_curr = get_ref(1)
        ref_price = get_ref(2)
        ref_b_ttm = get_ref(3)
        ref_b_prior = get_ref(4)
        ref_s_ttm = get_ref(5)
        ref_s_prior = get_ref(6)
        ref_elo, ref_emid, ref_ehigh = get_ref(7), get_ref(8), get_ref(9)
        ref_pec, ref_pel, ref_peh = get_ref(10), get_ref(11), get_ref(12)
        ref_cagr, ref_sector, ref_region = get_ref(13), get_ref(14), get_ref(15)
        ref_surprise = get_ref(16)
        ref_analysts = get_ref(17)
        ref_est_year = get_ref(18)
        ref_insider = get_ref(19)
        ref_name = get_ref(20)
        ref_mcap = get_ref(21)
        ref_elo_nxt = get_ref(22)
        ref_emid_nxt = get_ref(23)
        ref_ehigh_nxt = get_ref(24)
        ref_est_year_nxt = get_ref(25)
        ref_insider_count = get_ref(26)
        ref_avg_stake = get_ref(27)
        ref_perf_1y = get_ref(28)
        ref_eps_trend_1y = get_ref(29)
        ref_pe_trend_1y = get_ref(30)
        ref_implied_fwd = get_ref(31)
        ref_fcr = get_ref(32)
        ref_erg_plus = get_ref(33)
        ref_regime = get_ref(34)
        ref_regime_signal = get_ref(35)
        
        # Quality & Risk mapping (Columns 36, 37, 38, 39)
        ref_profit_margin = get_ref(36)
        ref_fcf_yield = get_ref(37)
        ref_de_ratio = get_ref(38)
        ref_roe = get_ref(39)

        dash_row, dash_col = 2, 15
        lists = {'Growth': 'Inputs!$A$2:$A$3', 'EPS': 'Inputs!$B$2:$B$3', 'PE': 'Inputs!$C$2:$C$4',
                 'Type': 'Inputs!$D$2:$D$3', 'FY': 'Inputs!$E$2:$E$3'}

        sheet.write(dash_row, dash_col, "EPS Type", fmt['dash_label'])
        sheet.write(dash_row + 1, dash_col, "Growth Basis", fmt['dash_label'])
        sheet.write(dash_row + 2, dash_col, "EPS Basis", fmt['dash_label'])
        sheet.write(dash_row + 3, dash_col, "P/E Mode", fmt['dash_label'])
        sheet.write(dash_row + 4, dash_col, "Target FY Period", fmt['dash_label'])

        sheet.data_validation(dash_row, dash_col + 1, dash_row, dash_col + 1,
                              {'validate': 'list', 'source': lists['Type']})
        sheet.write(dash_row, dash_col + 1, 'Street (Adjusted)', fmt['dash_input'])
        sheet.data_validation(dash_row + 1, dash_col + 1, dash_row + 1, dash_col + 1,
                              {'validate': 'list', 'source': lists['Growth']})
        sheet.write(dash_row + 1, dash_col + 1, 'Analyst Consensus', fmt['dash_input'])
        sheet.data_validation(dash_row + 2, dash_col + 1, dash_row + 2, dash_col + 1,
                              {'validate': 'list', 'source': lists['EPS']})
        sheet.write(dash_row + 2, dash_col + 1, 'TTM (Current)', fmt['dash_input'])
        sheet.data_validation(dash_row + 3, dash_col + 1, dash_row + 3, dash_col + 1,
                              {'validate': 'list', 'source': lists['PE']})
        sheet.write(dash_row + 3, dash_col + 1, 'Flexible (Hist)', fmt['dash_input'])
        sheet.data_validation(dash_row + 4, dash_col + 1, dash_row + 4, dash_col + 1,
                              {'validate': 'list', 'source': lists['FY']})
        sheet.write(dash_row + 4, dash_col + 1, 'Current', fmt['dash_input'])

        cell_type = xl_rowcol_to_cell(dash_row, dash_col + 1)
        cell_growth = xl_rowcol_to_cell(dash_row + 1, dash_col + 1)
        cell_eps = xl_rowcol_to_cell(dash_row + 2, dash_col + 1)
        cell_pe = xl_rowcol_to_cell(dash_row + 3, dash_col + 1)
        cell_fy = xl_rowcol_to_cell(dash_row + 4, dash_col + 1)
        addr_type_input = xl_rowcol_to_cell(dash_row, dash_col + 1, row_abs=True, col_abs=True)

        dyn_elo = f'IF({cell_fy}="Next", {ref_elo_nxt}, {ref_elo})'
        dyn_emid = f'IF({cell_fy}="Next", {ref_emid_nxt}, {ref_emid})'
        dyn_ehigh = f'IF({cell_fy}="Next", {ref_ehigh_nxt}, {ref_ehigh})'

        f_act_ttm = f'IF({cell_type}="Basic (GAAP)", {ref_b_ttm}, {ref_s_ttm})'
        f_calc_base = f'IF({cell_eps}="TTM (Current)", {f_act_ttm}, {dyn_elo})'
        f_calc_growth = f'IF({cell_growth}="Analyst Consensus", ({dyn_emid}-{f_act_ttm})/{f_act_ttm}, {ref_cagr})'
        f_pe_rt = f'{ref_price}/{f_act_ttm}'
        f_target_mid = f'({f_act_ttm}*(1+{f_calc_growth}))'
        f_pe_rt_bounded = f'MAX({MIN_HISTORICAL_PE_LOW}, {f_pe_rt})'
        f_calc_pemid = f'IF(ISNUMBER(SEARCH("Static",{cell_pe})), 30, IF(ISNUMBER(SEARCH("Custom",{cell_pe})), 20, {f_pe_rt_bounded}))'
        f_pe_down_step = f'IF(ISNUMBER(SEARCH("Static",{cell_pe})), 5, IF(ISNUMBER(SEARCH("Custom",{cell_pe})), 0, ({f_calc_pemid}-MIN({f_calc_pemid}-3, {ref_pel}))/3))'
        f_pe_up_step = f'IF(ISNUMBER(SEARCH("Static",{cell_pe})), 5, IF(ISNUMBER(SEARCH("Custom",{cell_pe})), 0, (MAX({f_calc_pemid}+3, {ref_peh})-{f_calc_pemid})/3))'
        f_calc_espstep = f'({f_target_mid}-{f_calc_base})/3'
        f_step_upper = f'({dyn_ehigh}-{f_target_mid})/3'
        pe_steps_offsets = [-3, -2, -1, 0, 1, 2, 3]


        def draw_grid_formulas(start_row, table_type):
            start_col = 1
            titles = {'PRICE': f'Implied Stock Price: {d["ticker"]}', 'UPSIDE': 'Implied Upside / Downside %',
                      'PEG': 'Implied PEG Ratio', 'HOLDEN': 'The Holden Score (Upside Efficiency)'}
            sheet.merge_range(start_row, start_col, start_row, start_col + 8, titles[table_type], fmt['title'])
            pe_row = start_row + 2
            sheet.merge_range(start_row + 1, start_col + 2, start_row + 1, start_col + 8, "P/E Multiple",
                              fmt['pe_label'])

            for j in range(7):
                offset = pe_steps_offsets[j]
                f_pe_cell_fmt = fmt['pe_val_mid'] if j == 3 else fmt['pe_val_base']
                if offset < 0:
                    form = f'={f_calc_pemid} + ({f_pe_down_step}*{offset})'
                elif offset > 0:
                    form = f'={f_calc_pemid} + ({f_pe_up_step}*{offset})'
                else:
                    form = f'={f_calc_pemid}'
                sheet.write_formula(pe_row, start_col + 2 + j, form, f_pe_cell_fmt)

            sheet.merge_range(start_row + 3, start_col, start_row + 9, start_col, "EPS", fmt['eps_label'])
            eps_col, eps_start_row = start_col + 1, start_row + 3

            for i in range(7):
                r = eps_start_row + i
                f_eps_cell_fmt = fmt['eps_mid'] if i == 3 else (fmt['eps_low'] if i == 0 else fmt['eps_val'])
                if i <= 3:
                    form = f'={f_calc_base} + ({f_calc_espstep}*{i})'
                else:
                    form = f'={f_target_mid} + ({f_step_upper}*{i - 3})'
                sheet.write_formula(r, eps_col, form, f_eps_cell_fmt)

            for i in range(7):
                r = eps_start_row + i
                eps_cell = xl_rowcol_to_cell(r, eps_col, col_abs=True)
                for j in range(7):
                    c = start_col + 2 + j
                    pe_cell = xl_rowcol_to_cell(pe_row, c, row_abs=True)
                    zone = 'center_zone' if abs(i - 3) == 0 and abs(j - 3) == 0 else (
                        'mid_zone' if abs(i - 3) <= 1 and abs(j - 3) <= 1 else 'outer_zone')
                    if table_type in ['UPSIDE', 'HOLDEN']: zone += '_pct'

                    if table_type == 'PRICE':
                        sheet.write_formula(r, c, f'={eps_cell}*{pe_cell}', fmt[zone])
                    elif table_type == 'UPSIDE':
                        sheet.write_formula(r, c, f'=({eps_cell}*{pe_cell} - {ref_price}) / {ref_price}', fmt[zone])
                    elif table_type == 'PEG':
                        growth_denom = f'(({eps_cell}-{f_act_ttm})/{f_act_ttm}*100)'
                        peg_f = f'=IF({growth_denom} < 0.1, "NM", {pe_cell}/{growth_denom})'
                        sheet.write_formula(r, c, peg_f, fmt['peg_nm'])
                    elif table_type == 'HOLDEN':
                        growth_denom = f'(({eps_cell}-{f_act_ttm})/{f_act_ttm}*100)'
                        upside_part = f'(({eps_cell}*{pe_cell} - {ref_price}) / {ref_price})'
                        peg_part = f'({pe_cell}/{growth_denom})'
                        form = f'=IF(OR({growth_denom}<0.1, {upside_part}<0), "NM", {upside_part}/{peg_part})'
                        sheet.write_formula(r, c, form, fmt[f'holden_{zone.split("_")[0]}'])

            if table_type == 'PEG':
                rng = f"{xl_rowcol_to_cell(start_row + 3, start_col + 2)}:{xl_rowcol_to_cell(start_row + 9, start_col + 8)}"
                for crit, val, f in [('<', 0.75, 'peg_deep_green'), ('between', 0.75, 'peg_green'),
                                     ('between', 1.0, 'peg_yellow'),
                                     ('between', 1.5, 'peg_orange'), ('between', 2.0, 'peg_red'),
                                     ('>', 3.0, 'peg_black')]:
                    props = {'type': 'cell', 'criteria': crit, 'format': fmt[f]}
                    if crit == 'between':
                        props.update({'minimum': val, 'maximum':
                            {'peg_green': 1.0, 'peg_yellow': 1.5, 'peg_orange': 2.0, 'peg_red': 3.0}[f]})
                    else:
                        props['value'] = val
                    sheet.conditional_format(rng, props)


        draw_grid_formulas(1, 'PRICE')
        draw_grid_formulas(14, 'UPSIDE')
        draw_grid_formulas(27, 'PEG')
        draw_grid_formulas(40, 'HOLDEN')

        # --- SUMMARY STATISTICS ---
        stats_col, r = 11, 1
        sheet.write(r, stats_col, "Summary Statistics", fmt['stat_head'])
        r += 1
        sheet.write(r, stats_col, "Market Data", fmt['stat_subhead'])
        r += 1
        sheet.write(r, stats_col, "Name", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_name}', fmt['stat_val_txt'])
        r += 1
        sheet.write(r, stats_col, "Sector", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_sector}', fmt['stat_val_txt'])
        r += 1
        sheet.write(r, stats_col, "Market Cap", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_mcap}', fmt['stat_val_mcap'])
        r += 1
        sheet.write(r, stats_col, "Region", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_region}', fmt['stat_val_txt'])
        r += 1
        sheet.write(r, stats_col, "Current Price", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_price}', fmt['stat_val'])
        addr_price = xl_rowcol_to_cell(r, stats_col + 1)

        r += 2
        sheet.write(r, stats_col, "Earnings Basis", fmt['stat_subhead'])
        r += 1
        sheet.write(r, stats_col, "GAAP EPS (Diluted)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_b_ttm}', fmt['stat_val_blue'])
        addr_basic = xl_rowcol_to_cell(r, stats_col + 1)
        r += 1
        sheet.write(r, stats_col, "Street EPS (Adj)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_s_ttm}', fmt['stat_val_blue'])
        addr_street = xl_rowcol_to_cell(r, stats_col + 1)

        r += 2
        r_forecast_start = r

        # Current FY Block
        sheet.write(r, stats_col, "Analyst Forecasts", fmt['stat_subhead'])
        r += 1
        sheet.write(r, stats_col, "Target FY Period", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_est_year}', fmt['stat_val_txt'])
        r += 1
        sheet.write(r, stats_col, "EPS Low (Target FY)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_elo}', fmt['stat_val_blue'])
        r += 1
        sheet.write(r, stats_col, "EPS Consensus", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_emid}', fmt['stat_val_purple'])
        sheet.write_formula(r, stats_col + 2, f'=({ref_emid}-{f_act_ttm})/{f_act_ttm}', fmt['est_growth_base'])
        r += 1
        sheet.write(r, stats_col, "EPS High (Target FY)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_ehigh}', fmt['stat_val'])
        r += 1
        sheet.write(r, stats_col, "Growth Diagnosis", fmt['stat_label'])
        f_diag = f'=IF({addr_type_input}="Basic (GAAP)", IF({ref_b_ttm}<{ref_b_prior}, "⚠️ Cyclical Rebound", "✔ Organic"), IF({ref_s_ttm}<{ref_s_prior}, "⚠️ Cyclical Rebound", "✔ Organic"))'
        sheet.write_formula(r, stats_col + 1, f_diag, fmt['diag_ok'])

        # Next FY Block
        nxt_col = stats_col + 4
        r_nxt = r_forecast_start
        sheet.write(r_nxt, nxt_col, "Analyst Forecasts (Next FY)", fmt['stat_subhead'])
        r_nxt += 1
        sheet.write(r_nxt, nxt_col, "Target FY Period", fmt['stat_label'])
        sheet.write_formula(r_nxt, nxt_col + 1, f'={ref_est_year_nxt}', fmt['stat_val_txt'])
        r_nxt += 1
        sheet.write(r_nxt, nxt_col, "EPS Low (Target FY)", fmt['stat_label'])
        sheet.write_formula(r_nxt, nxt_col + 1, f'={ref_elo_nxt}', fmt['stat_val_blue'])
        r_nxt += 1
        sheet.write(r_nxt, nxt_col, "EPS Consensus", fmt['stat_label'])
        sheet.write_formula(r_nxt, nxt_col + 1, f'={ref_emid_nxt}', fmt['stat_val_purple'])
        sheet.write_formula(r_nxt, nxt_col + 2, f'=({ref_emid_nxt}-{f_act_ttm})/{f_act_ttm}', fmt['est_growth_base'])
        r_nxt += 1
        sheet.write(r_nxt, nxt_col, "EPS High (Target FY)", fmt['stat_label'])
        sheet.write_formula(r_nxt, nxt_col + 1, f'={ref_ehigh_nxt}', fmt['stat_val'])

        r += 1
        sheet.write(r, stats_col, "Credibility (Surprise)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_surprise}/100', fmt['stat_val_score'])
        r += 1
        sheet.write(r, stats_col, "Estimate Dispersion", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'=({dyn_ehigh}-{dyn_elo})/{dyn_emid}', fmt['stat_val_score'])
        r += 1
        sheet.write(r, stats_col, "Analyst Coverage (#)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_analysts}', fmt['stat_val_txt'])

        # --- QUALITY & RISK PROFILE ---
        r += 2
        sheet.write(r, stats_col, "Quality & Risk Profile", fmt['stat_subhead'])
        
        # Net Profit Margin
        r += 1
        sheet.write(r, stats_col, "Net Profit Margin", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_profit_margin}', fmt['stat_val_score'])
        cell_pm = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.conditional_format(cell_pm, {'type': 'cell', 'criteria': '>=', 'value': 0.30, 'format': fmt['tier_dkgreen']})
        sheet.conditional_format(cell_pm, {'type': 'cell', 'criteria': 'between', 'minimum': 0.20, 'maximum': 0.2999, 'format': fmt['tier_green']})
        sheet.conditional_format(cell_pm, {'type': 'cell', 'criteria': 'between', 'minimum': 0.10, 'maximum': 0.1999, 'format': fmt['tier_yellow']})
        sheet.conditional_format(cell_pm, {'type': 'cell', 'criteria': 'between', 'minimum': 0.05, 'maximum': 0.0999, 'format': fmt['tier_red']})
        sheet.conditional_format(cell_pm, {'type': 'cell', 'criteria': '<', 'value': 0.05, 'format': fmt['tier_black']})

        # FCF Yield
        r += 1
        sheet.write(r, stats_col, "Free Cash Flow Yield", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_fcf_yield}', fmt['stat_val_score'])
        cell_fcf = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.conditional_format(cell_fcf, {'type': 'cell', 'criteria': '>=', 'value': 0.08, 'format': fmt['tier_dkgreen']})
        sheet.conditional_format(cell_fcf, {'type': 'cell', 'criteria': 'between', 'minimum': 0.05, 'maximum': 0.0799, 'format': fmt['tier_green']})
        sheet.conditional_format(cell_fcf, {'type': 'cell', 'criteria': 'between', 'minimum': 0.025, 'maximum': 0.0499, 'format': fmt['tier_yellow']})
        sheet.conditional_format(cell_fcf, {'type': 'cell', 'criteria': 'between', 'minimum': 0.0, 'maximum': 0.0249, 'format': fmt['tier_red']})
        sheet.conditional_format(cell_fcf, {'type': 'cell', 'criteria': '<', 'value': 0.0, 'format': fmt['tier_black']})
        
        # Debt-To-Equity
        r += 1
        sheet.write(r, stats_col, "Debt-to-Equity", fmt['stat_label'])
        f_de = f'=IF({ref_de_ratio}<0, "N/A", {ref_de_ratio})'
        sheet.write_formula(r, stats_col + 1, f_de, fmt['stat_val_pe'])
        cell_de = xl_rowcol_to_cell(r, stats_col + 1)
        # Using format mapping
        sheet.conditional_format(cell_de, {'type': 'cell', 'criteria': 'between', 'minimum': 0.0, 'maximum': 0.499, 'format': fmt['tier_de_dkgreen']})
        sheet.conditional_format(cell_de, {'type': 'cell', 'criteria': 'between', 'minimum': 0.5, 'maximum': 0.999, 'format': fmt['tier_de_green']})
        sheet.conditional_format(cell_de, {'type': 'cell', 'criteria': 'between', 'minimum': 1.0, 'maximum': 1.999, 'format': fmt['tier_de_yellow']})
        sheet.conditional_format(cell_de, {'type': 'cell', 'criteria': 'between', 'minimum': 2.0, 'maximum': 3.999, 'format': fmt['tier_de_red']})
        sheet.conditional_format(cell_de, {'type': 'cell', 'criteria': '>=', 'value': 4.0, 'format': fmt['tier_de_black']})
        
        # Return on Equity (ROE)
        r += 1
        sheet.write(r, stats_col, "Return on Equity (ROE)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_roe}', fmt['stat_val_score'])
        cell_roe = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.conditional_format(cell_roe, {'type': 'cell', 'criteria': '>=', 'value': 0.25, 'format': fmt['tier_dkgreen']})
        sheet.conditional_format(cell_roe, {'type': 'cell', 'criteria': 'between', 'minimum': 0.15, 'maximum': 0.2499, 'format': fmt['tier_green']})
        sheet.conditional_format(cell_roe, {'type': 'cell', 'criteria': 'between', 'minimum': 0.08, 'maximum': 0.1499, 'format': fmt['tier_yellow']})
        sheet.conditional_format(cell_roe, {'type': 'cell', 'criteria': 'between', 'minimum': 0.0, 'maximum': 0.0799, 'format': fmt['tier_red']})
        sheet.conditional_format(cell_roe, {'type': 'cell', 'criteria': '<', 'value': 0.0, 'format': fmt['tier_black']})

        r += 2
        sheet.write(r, stats_col, "Valuation Logic", fmt['stat_subhead'])
        r += 1
        sheet.write(r, stats_col, "Implied P/E (Active)", fmt['stat_label'])
        addr_active_pe = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.write_formula(r, stats_col + 1,
                            f'={addr_price} / IF({addr_type_input}="Basic (GAAP)", {addr_basic}, {addr_street})',
                            fmt['stat_val_pe_grey'])
        r += 1
        sheet.write(r, stats_col, "Forward P/E (Est.)", fmt['stat_label'])
        addr_fwd_pe = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.write_formula(r, stats_col + 1, f'={addr_price} / {dyn_emid}', fmt['stat_val_pe'])

        r += 2
        sheet.write(r, stats_col, "Holden Resilience", fmt['stat_subhead'])
        r += 1
        sheet.write(r, stats_col, "Allowed Safety Cushion", fmt['stat_label'])
        addr_cushion = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.write_formula(r, stats_col + 1, f'=({addr_active_pe} - {addr_fwd_pe}) / {addr_active_pe}',
                            fmt['cushion_val'])
        r += 1
        sheet.write(r, stats_col, "Hist. Downside Risk", fmt['stat_label'])
        addr_risk = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.write_formula(r, stats_col + 1, f'=MAX(0, ({addr_active_pe} - {ref_pel}) / {addr_active_pe})',
                            fmt['risk_val'])
        r += 1
        sheet.write(r, stats_col, "Resilience Ratio", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'=IF({addr_risk}<=0, 99, {addr_cushion} / {addr_risk})',
                            fmt['resilience_good'])

        r += 2
        sheet.write(r, stats_col, "Holden Score (Base)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, '=G47', fmt['stat_val_score'])
        r += 1
        sheet.write(r, stats_col, "Realizable Upside %", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, '=IF(G34="NM", "NM", G21/MAX(1, G34))', fmt['stat_val_real'])

        # --- HISTORICAL TRENDS (1Y) ---
        r += 2
        sheet.write(r, stats_col, "Historical Trends (1Y)", fmt['stat_subhead'])
        r += 1
        sheet.write(r, stats_col, "1Y Price Performance", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_perf_1y}', fmt['stat_val_score'])
        trend_cell = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.conditional_format(trend_cell, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt['diag_warn']})
        sheet.conditional_format(trend_cell, {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt['diag_ok']})
        r += 1
        sheet.write(r, stats_col, "1Y EPS Trend", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_eps_trend_1y}', fmt['stat_val_score'])
        trend_cell = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.conditional_format(trend_cell, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt['diag_warn']})
        sheet.conditional_format(trend_cell, {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt['diag_ok']})
        r += 1
        sheet.write(r, stats_col, "1Y P/E Trend", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_pe_trend_1y}', fmt['stat_val_score'])
        trend_cell = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.conditional_format(trend_cell, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt['diag_warn']})
        sheet.conditional_format(trend_cell, {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt['diag_ok']})
        r += 1
        sheet.write(r, stats_col, "1Y Earnings-Return Gap (ERG)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'=IF({ref_eps_trend_1y}>0, {ref_eps_trend_1y} - {ref_perf_1y}, "N/A (EPS < 0)")', fmt['stat_val_score'])
        ufa_cell = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.conditional_format(ufa_cell, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt['diag_warn']})
        sheet.conditional_format(ufa_cell, {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt['diag_ok']})

        # --- FORWARD-CONFIRMED ERG ---
        r += 2
        sheet.write(r, stats_col, "Forward-Confirmed ERG", fmt['stat_subhead'])
        r += 1
        sheet.write(r, stats_col, "Implied FWD Growth", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_implied_fwd}', fmt['stat_val_score'])
        fwd_cell = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.conditional_format(fwd_cell, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt['diag_warn']})
        sheet.conditional_format(fwd_cell, {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt['diag_ok']})
        r += 1
        sheet.write(r, stats_col, "Growth Confirm. (FCR)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_fcr}', fmt['stat_val_fcr'])
        fcr_cell = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.conditional_format(fcr_cell, {'type': 'cell', 'criteria': '<', 'value': 0.3, 'format': fmt['diag_warn']})
        sheet.conditional_format(fcr_cell, {'type': 'cell', 'criteria': '>=', 'value': 0.7, 'format': fmt['diag_ok']})
        r += 1
        sheet.write(r, stats_col, "ERG+ (Confirmed)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_erg_plus}', fmt['stat_val_score'])
        erg_plus_cell = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.conditional_format(erg_plus_cell, {'type': 'cell', 'criteria': '<=', 'value': 0, 'format': fmt['diag_warn']})
        sheet.conditional_format(erg_plus_cell, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt['diag_ok']})
        r += 1
        sheet.write(r, stats_col, "Regime", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_regime}', fmt['stat_val_txt'])
        r += 1
        sheet.write(r, stats_col, "Implied Signal", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_regime_signal}', fmt['stat_val_txt'])

        # --- INSIDER ACTIVITY & CONVICTION SCORING ENGINE ---
        r += 2
        sheet.write(r, stats_col, "Insider Activity (L6M)", fmt['stat_subhead'])
        r += 1
        sheet.write(r, stats_col, "Net Buying ($)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_insider}', fmt['insider_neutral'])
        insider_cell = xl_rowcol_to_cell(r, stats_col + 1)
        sheet.conditional_format(insider_cell,
                                 {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt['insider_pos']})
        sheet.conditional_format(insider_cell,
                                 {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt['insider_neg']})
        r += 1
        sheet.write(r, stats_col, "Unique Buyers (#)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_insider_count}', fmt['stat_val_int'])
        r += 1
        sheet.write(r, stats_col, "Avg Stake Inc. (%)", fmt['stat_label'])
        sheet.write_formula(r, stats_col + 1, f'={ref_avg_stake}/100', fmt['stat_val_score'])

        # Conviction Scoring Formulas (Lone Wolf Fix)
        r += 1
        sheet.write(r, stats_col, "Conviction Score (0-10)", fmt['stat_label'])

        # Pillar 1: Materiality (Max 2)
        p1 = f'IF({ref_insider}<=0, 0, IF(OR({ref_insider}/MAX(1,{ref_mcap})>{INSIDER_PCT_LARGE}, {ref_insider}>={INSIDER_DOLLAR_LARGE}), 2, IF(OR({ref_insider}/MAX(1,{ref_mcap})>{INSIDER_PCT_MODERATE}, {ref_insider}>={INSIDER_DOLLAR_MODERATE}), 1, 0)))'
        # Pillar 2: Breadth (Max 4)
        p2 = f'IF({ref_insider_count}>=4, 4, {ref_insider_count})'
        # Pillar 3: Depth (Max 4) - Capped at 1 if Unique Buyers == 1
        p3 = f'IF({ref_insider_count}=0, 0, IF({ref_insider_count}=1, IF({ref_avg_stake}>={INSIDER_STAKE_PCT_FOR_SCORE_1}, 1, 0), IF({ref_avg_stake}>={INSIDER_STAKE_PCT_FOR_SCORE_4}, 4, IF({ref_avg_stake}>={INSIDER_STAKE_PCT_FOR_SCORE_3}, 3, IF({ref_avg_stake}>={INSIDER_STAKE_PCT_FOR_SCORE_2}, 2, IF({ref_avg_stake}>={INSIDER_STAKE_PCT_FOR_SCORE_1}, 1, 0))))))'

        score_formula = f'=({p1}) + ({p2}) + ({p3})'
        sheet.write_formula(r, stats_col + 1, score_formula, fmt['stat_val_score_int'])

        foot_row = r + 2
        sheet.write(foot_row, stats_col, "Holden Resilience Interpretation:", fmt['legend_bold'])
        sheet.write(foot_row + 1, stats_col, "> 1.0: Growth cushion covers historical downside (Safe)",
                    fmt['legend_norm'])

        sheet.set_column(0, 0, 2)
        sheet.set_column(1, 1, 6)
        sheet.set_column(2, 2, 12)
        sheet.set_column(3, 9, 10)
        sheet.set_column(11, 11, 22)
        sheet.set_column(12, 12, 16)
        sheet.set_column(13, 13, 12)
        sheet.set_column(14, 14, 2)
        sheet.set_column(15, 15, 22)
        sheet.set_column(16, 16, 18)

    # 2. CREATE COMPARISON SHEET
    print("  Generating Comparison Sheet...")
    ws_comp = workbook.add_worksheet("Comparison")
    fmt_comp_h = fmt['comp_head']

    cols = [
        "Ticker", "Company Name", "Sector", "Market Cap",
        "Price", "Target Price", "Implied Upside",
        "Current P/E (Adj)", "Forward P/E", "PEG Ratio",
        "Holden Score", "Safety Cushion", "Resilience Ratio",
        "Insider Net L6M ($)", "Unique Buyers", "Avg Stake Inc (%)",
        "Conviction Score (0-10)", "Growth Diagnosis",
        "1Y Perf", "1Y EPS \u0394", "1Y P/E \u0394", "1Y ERG Score",
        "Impl. FWD Growth", "FCR", "ERG+", "Regime", "Signal",
        "Net Profit Margin", "FCF Yield", "Debt/Equity", "ROE"
    ]
    ws_comp.write_row(0, 0, cols, fmt_comp_h)
    comp_data = []

    for d in ALL_DATA:
        tick, name, sector, mcap, price = d['ticker'], d['company_name'], d['sector'], d['market_cap'], d['price']
        street_ttm, street_prior = d['eps_street_ttm'], d['eps_street_prior']
        basic_ttm, basic_prior = d['eps_basic_ttm'], d['eps_basic_prior']

        eps_ttm = street_ttm if street_ttm != 0 else basic_ttm
        if eps_ttm == 0: eps_ttm = 0.01
        eps_fwd = d['eps_mid'] if d['eps_mid'] != 0 else 0.01

        pe_low_hist = d['pe_low_hist']
        pe_curr = price / eps_ttm
        target_price = pe_curr * eps_fwd
        upside = (target_price - price) / price if price != 0 else 0
        pe_fwd = price / eps_fwd
        growth = (eps_fwd - eps_ttm) / eps_ttm
        peg = pe_curr / (growth * 100) if growth > 0.001 else 999.0
        holden_score = upside / peg if (peg > 0 and upside > 0 and peg != 999.0) else 0
        cushion = (pe_curr - pe_fwd) / pe_curr if pe_curr != 0 else 0
        risk = (pe_curr - pe_low_hist) / pe_curr if pe_curr != 0 else 0
        resilience = cushion / risk if risk > 0 else 99.0

        insider = d['insider_net']
        insider_count = d['insider_buy_count']
        avg_stake = d['insider_avg_stake_inc']

        # Calculate Conviction Score (Lone Wolf Fix applied here)
        p1 = 0
        if insider > 0:
            m = max(1, mcap)
            if (insider / m > INSIDER_PCT_LARGE) or (insider >= INSIDER_DOLLAR_LARGE):
                p1 = 2
            elif (insider / m > INSIDER_PCT_MODERATE) or (insider >= INSIDER_DOLLAR_MODERATE):
                p1 = 1
        p2 = min(4, insider_count)
        p3 = 0
        if insider_count == 1:
            if avg_stake >= INSIDER_STAKE_PCT_FOR_SCORE_1: p3 = 1
        elif insider_count > 1:
            if avg_stake >= INSIDER_STAKE_PCT_FOR_SCORE_4:
                p3 = 4
            elif avg_stake >= INSIDER_STAKE_PCT_FOR_SCORE_3:
                p3 = 3
            elif avg_stake >= INSIDER_STAKE_PCT_FOR_SCORE_2:
                p3 = 2
            elif avg_stake >= INSIDER_STAKE_PCT_FOR_SCORE_1:
                p3 = 1

        conviction_score = p1 + p2 + p3
        growth_diag = "⚠️ Cyclical Rebound" if (
            street_ttm < street_prior if street_ttm != 0 else basic_ttm < basic_prior) else "✔ Organic"

        eps_trend_1y = d['eps_trend_1y']
        pe_trend_1y = d['pe_trend_1y']
        perf_1y = d['perf_1y']
        erg_score = (eps_trend_1y - perf_1y) if eps_trend_1y > 0 else "N/A"
        implied_fwd_g = d['implied_fwd_growth']
        fcr_val = d['fcr']
        erg_plus_val = d['erg_plus']
        regime_val = d['regime']
        signal_val = d['regime_signal']

        comp_data.append([tick, name, sector, mcap, price, target_price, upside, pe_curr, pe_fwd, peg,
                          holden_score, cushion, resilience, insider, insider_count, avg_stake / 100, conviction_score,
                          growth_diag, perf_1y, eps_trend_1y, pe_trend_1y, erg_score,
                          implied_fwd_g, fcr_val, erg_plus_val, regime_val, signal_val,
                          d['profit_margin'], d['fcf_yield'], d['de_ratio'], d['roe']])

    for i, row in enumerate(comp_data):
        ws_comp.write(i + 1, 0, row[0], fmt['comp_ticker'])
        ws_comp.write(i + 1, 1, row[1], fmt['comp_txt'])
        ws_comp.write(i + 1, 2, row[2], fmt['comp_txt'])
        ws_comp.write(i + 1, 3, row[3], fmt['comp_mcap'])
        ws_comp.write(i + 1, 4, row[4], fmt['comp_num'])
        ws_comp.write(i + 1, 5, row[5], fmt['comp_num'])
        ws_comp.write(i + 1, 6, row[6], fmt['comp_pct'])
        ws_comp.write(i + 1, 7, row[7], fmt['comp_pe'])
        ws_comp.write(i + 1, 8, row[8], fmt['comp_pe'])
        ws_comp.write(i + 1, 9, row[9], fmt['comp_num'])
        ws_comp.write(i + 1, 10, row[10], fmt['comp_pct'])
        ws_comp.write(i + 1, 11, row[11], fmt['comp_pct'])
        ws_comp.write(i + 1, 12, row[12], fmt['comp_num'])
        ws_comp.write(i + 1, 13, row[13], fmt['comp_dollar'])
        ws_comp.write(i + 1, 14, row[14], fmt['comp_int'])
        ws_comp.write(i + 1, 15, row[15], fmt['stat_val_score'])
        ws_comp.write(i + 1, 16, row[16], fmt['comp_int'])
        ws_comp.write(i + 1, 17, row[17], fmt['comp_txt'])
        ws_comp.write(i + 1, 18, row[18], fmt['comp_pct'])
        ws_comp.write(i + 1, 19, row[19], fmt['comp_pct'])
        ws_comp.write(i + 1, 20, row[20], fmt['comp_pct'])
        if isinstance(row[21], str):
            ws_comp.write(i + 1, 21, row[21], fmt['comp_txt'])
        else:
            ws_comp.write(i + 1, 21, row[21], fmt['comp_pct'])
        ws_comp.write(i + 1, 22, row[22], fmt['comp_pct'])
        ws_comp.write(i + 1, 23, row[23], fmt['comp_num'])
        ws_comp.write(i + 1, 24, row[24], fmt['comp_pct'])
        ws_comp.write(i + 1, 25, row[25], fmt['comp_txt'])
        ws_comp.write(i + 1, 26, row[26], fmt['comp_txt'])
        ws_comp.write(i + 1, 27, row[27], fmt['comp_pct'])
        ws_comp.write(i + 1, 28, row[28], fmt['comp_pct'])
        val_de = row[29]
        if val_de < 0:
            ws_comp.write(i + 1, 29, "N/A", fmt['comp_txt'])
        else:
            ws_comp.write(i + 1, 29, val_de, fmt['comp_num'])
        ws_comp.write(i + 1, 30, row[30], fmt['comp_pct'])

    if comp_data:
        ws_comp.add_table(0, 0, len(comp_data), len(cols) - 1, {
            'columns': [{'header': c} for c in cols],
            'style': 'TableStyleMedium2',
            'name': 'ValuationComparison'
        })

    ws_comp.set_column(0, 0, 10)
    ws_comp.set_column(1, 1, 25)
    ws_comp.set_column(2, 2, 22)
    ws_comp.set_column(3, 3, 16)
    ws_comp.set_column(4, 16, 14)
    ws_comp.set_column(17, 17, 20)
    ws_comp.set_column(18, 24, 14)
    ws_comp.set_column(25, 26, 22)
    ws_comp.set_column(27, 28, 14)
    ws_comp.set_column(29, 30, 14)

    # 3. CREATE INPUTS SHEET
    print("  Generating Inputs Sheet...")
    ws_inputs = workbook.add_worksheet("Inputs")
    try:
        ws_inputs.set_tab_color('#808080')
    except:
        pass

    lists = {'Growth': ['Analyst Consensus', 'Historical CAGR'], 'EPS': ['TTM (Current)', 'Analyst Low Est'],
             'PE': ['Flexible (Hist)', 'Static (15-45x)', 'Custom (20x)'],
             'Type': ['Basic (GAAP)', 'Street (Adjusted)'],
             'FY': ['Current', 'Next']}

    ws_inputs.write(0, 0, "List: Growth", workbook.add_format({'bold': True}))
    ws_inputs.write_column(1, 0, lists['Growth'])
    ws_inputs.write(0, 1, "List: EPS", workbook.add_format({'bold': True}))
    ws_inputs.write_column(1, 1, lists['EPS'])
    ws_inputs.write(0, 2, "List: PE", workbook.add_format({'bold': True}))
    ws_inputs.write_column(1, 2, lists['PE'])
    ws_inputs.write(0, 3, "List: Type", workbook.add_format({'bold': True}))
    ws_inputs.write_column(1, 3, lists['Type'])
    ws_inputs.write(0, 4, "List: FY", workbook.add_format({'bold': True}))
    ws_inputs.write_column(1, 4, lists['FY'])

    headers = ['Ticker', 'Currency', 'Price', 'EPS Basic TTM', 'EPS Basic Prior', 'EPS Street TTM', 'EPS Street Prior',
               'EPS Low', 'EPS Mid', 'EPS High', 'PE Current', 'PE Low Hist', 'PE High Hist', 'CAGR 3Y', 'Sector',
               'Region',
               'Surprise Avg 4Q', 'Analyst Count', 'Est Target FY', 'Insider Net L6M', 'Company Name', 'Market Cap',
               'EPS Low Nxt', 'EPS Mid Nxt', 'EPS High Nxt', 'Est Target FY Nxt', 'Unique Buyers Count',
               'Avg Stake Increase %', 'Perf 1Y', 'EPS Trend 1Y', 'PE Trend 1Y',
               'Implied FWD Growth', 'FCR', 'ERG+', 'Regime', 'Regime Signal',
               'Profit Margin', 'FCF Yield', 'Debt to Equity', 'ROE']

    fmt_head = workbook.add_format({'bold': True, 'bottom': 1})
    ws_inputs.write_row(data_start_row, 0, headers, fmt_head)

    for i, d in enumerate(ALL_DATA):
        r = data_start_row + 1 + i
        row_data = [d['ticker'], d['currency'], d['price'], d['eps_basic_ttm'], d.get('eps_basic_prior', 0),
                    d.get('eps_street_ttm', 0), d.get('eps_street_prior', 0), d['eps_low'], d['eps_mid'], d['eps_high'],
                    d['pe_current'], d['pe_low_hist'], d['pe_high_hist'], d['cagr_3y'], d['sector'], d['region'],
                    d.get('surprise_avg', 0), d.get('analysts', 0), d['est_year_str'], d['insider_net'],
                    d['company_name'], d['market_cap'], d['eps_low_nxt'], d['eps_mid_nxt'], d['eps_high_nxt'],
                    d['est_year_str_nxt'], d['insider_buy_count'], d['insider_avg_stake_inc'],
                    d['perf_1y'], d['eps_trend_1y'], d['pe_trend_1y'],
                    d['implied_fwd_growth'], d['fcr'], d['erg_plus'], d['regime'], d['regime_signal'],
                    d['profit_margin'], d['fcf_yield'], d['de_ratio'], d['roe']]
        ws_inputs.write_row(r, 0, row_data)

    writer.close()
    print(f"\nSaved successfully to: {FILENAME}")