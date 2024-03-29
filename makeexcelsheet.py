import xlsxwriter

def fillcell(worksheet,coveragedata,plaincellformat,position,value):
    if value in coveragedata:
        worksheet.write(position,coveragedata[value],plaincellformat)
    else:
        worksheet.write(position,"N/A",plaincellformat)

def makeexcelsheet(workbook,worksheet,coveragedata):
    #Make name of worksheet
    worksheet = workbook.add_worksheet(worksheet)

    worksheet.set_row(0, 30)
    worksheet.set_column('A:H', 8.43)
    worksheet.set_column('K:K', 21.57)

    #Format of cells
    mergedheadersformat = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'left',
        'valign': 'vcenter'})
    kcolumnheaderformat = workbook.add_format({
        'bold': 1,
        'align': 'left',
        'valign': 'vcenter'})
    italicgenelabelcellformat = workbook.add_format({
        'italic': 1,
        'align': 'left',
        'valign': 'vcenter',
        'border': 1})
    bolditalicgenelabelcellformat = workbook.add_format({
        'italic': 1,
        'bold': 1,
        'align': 'left',
        'valign': 'vcenter',
        'border': 1})
    plaincellformat = workbook.add_format({
        'align': 'left',
        'valign': 'vcenter',
        'border': 1})

    #Make all the merged cells
    worksheet.merge_range('A1:B1','RAS/TK Signalling',mergedheadersformat)
    worksheet.merge_range('C1:D1','Other signalling\npathway genes',mergedheadersformat)
    worksheet.merge_range('E1:F1','Cohesin pathway',mergedheadersformat)
    worksheet.merge_range('G1:H1','Splicing pathway',mergedheadersformat)
    
    #These remove GNAS
    worksheet.merge_range('C5:D6','DNA/chromatin\nmodification',mergedheadersformat)
    
    #These are old versions prior to removing GNAS
    #worksheet.merge_range('C6:D7','DNA/chromatin\nmodification',mergedheadersformat)
    
    worksheet.merge_range('E6:F7','Transcription\nfactors',mergedheadersformat)
    worksheet.merge_range('G6:H7','Other',mergedheadersformat)

    #Add labels in single cells
    worksheet.write('K1','Hotspot coverage (100x)',kcolumnheaderformat)
    worksheet.write('A2','BRAF',italicgenelabelcellformat)
    worksheet.write('A3','CALR',italicgenelabelcellformat)
    worksheet.write('A4','CBL',italicgenelabelcellformat)
    worksheet.write('A5','CBLB',italicgenelabelcellformat)
    worksheet.write('A6','CSF3R',italicgenelabelcellformat)
    worksheet.write('A7','FLT3',italicgenelabelcellformat)
    worksheet.write('A8','HRAS',italicgenelabelcellformat)
    worksheet.write('A9','JAK2',italicgenelabelcellformat)
    worksheet.write('A10','JAK3',italicgenelabelcellformat)
    worksheet.write('A11','KIT',italicgenelabelcellformat)
    worksheet.write('A12','KRAS',italicgenelabelcellformat)
    worksheet.write('A13','MPL',italicgenelabelcellformat)
    worksheet.write('A14','NRAS',italicgenelabelcellformat)
    worksheet.write('A15','PDGFRA',italicgenelabelcellformat)
    worksheet.write('A16','PTPN11',italicgenelabelcellformat)
    
    #These remove GNAS
    worksheet.write('C2','MYD88',italicgenelabelcellformat)
    worksheet.write('C3','NOTCH1',italicgenelabelcellformat)
    worksheet.write('C4','PTEN',italicgenelabelcellformat)
    worksheet.write('C7','ASXL1',italicgenelabelcellformat)
    worksheet.write('C8','ATRX',italicgenelabelcellformat)
    worksheet.write('C9','BCOR',bolditalicgenelabelcellformat)
    worksheet.write('C10','BCORL1',bolditalicgenelabelcellformat)
    worksheet.write('C11','DNMT3A',bolditalicgenelabelcellformat)
    worksheet.write('C12','EZH2',bolditalicgenelabelcellformat)
    worksheet.write('C13','GATA1',italicgenelabelcellformat)
    worksheet.write('C14','IDH1',italicgenelabelcellformat)
    worksheet.write('C15','IDH2',italicgenelabelcellformat)
    worksheet.write('C16','KMT2A',italicgenelabelcellformat)
    worksheet.write('C17','PHF6',bolditalicgenelabelcellformat)
    worksheet.write('C18','TET2',italicgenelabelcellformat)
    
    #These are old versions prior to removing GNAS
    #worksheet.write('C2','GNAS',italicgenelabelcellformat)
    #worksheet.write('C3','MYD88',italicgenelabelcellformat)
    #worksheet.write('C4','NOTCH1',italicgenelabelcellformat)
    #worksheet.write('C5','PTEN',italicgenelabelcellformat)
    #worksheet.write('C8','ASXL1',italicgenelabelcellformat)
    #worksheet.write('C9','ATRX',italicgenelabelcellformat)
    #worksheet.write('C10','BCOR',bolditalicgenelabelcellformat)
    #worksheet.write('C11','BCORL1',bolditalicgenelabelcellformat)
    #worksheet.write('C12','DNMT3A',bolditalicgenelabelcellformat)
    #worksheet.write('C13','EZH2',bolditalicgenelabelcellformat)
    #worksheet.write('C14','GATA1',italicgenelabelcellformat)
    #worksheet.write('C15','IDH1',italicgenelabelcellformat)
    #worksheet.write('C16','IDH2',italicgenelabelcellformat)
    #worksheet.write('C17','KMT2A',italicgenelabelcellformat)
    #worksheet.write('C18','PHF6',bolditalicgenelabelcellformat)
    #worksheet.write('C19','TET2',italicgenelabelcellformat)
    
    worksheet.write('E2','RAD21',bolditalicgenelabelcellformat)
    worksheet.write('E3','SMC1A',italicgenelabelcellformat)
    worksheet.write('E4','SMC3',italicgenelabelcellformat)
    worksheet.write('E5','STAG2',bolditalicgenelabelcellformat)
    worksheet.write('E8','CUX1',bolditalicgenelabelcellformat)
    worksheet.write('E9','ETV6',bolditalicgenelabelcellformat)
    worksheet.write('E10','GATA2',italicgenelabelcellformat)
    worksheet.write('E11','IKZF1',bolditalicgenelabelcellformat)
    worksheet.write('E12','RUNX1',bolditalicgenelabelcellformat)
    worksheet.write('E13','WT1',italicgenelabelcellformat)
    worksheet.write('G2','SF3B1',italicgenelabelcellformat)
    worksheet.write('G3','SRSF2',italicgenelabelcellformat)
    worksheet.write('G4','U2AF1',italicgenelabelcellformat)
    worksheet.write('G5','ZRSR2',bolditalicgenelabelcellformat)
    worksheet.write('G8','FBXW7',italicgenelabelcellformat)
    worksheet.write('G9','NPM1',italicgenelabelcellformat)
    worksheet.write('G10','SETBP1',italicgenelabelcellformat)
    worksheet.write('G11','TP53',italicgenelabelcellformat)
    worksheet.write('K2','NRAShotspot12&13',plaincellformat)
    worksheet.write('K3','DNMT3AhotspotR882',plaincellformat)
    worksheet.write('K4','SF3B1hotspotK700',plaincellformat)
    worksheet.write('K5','SF3B1hotspotH662K666',plaincellformat)
    worksheet.write('K6','SF3B1hotspotE622N626',plaincellformat)
    worksheet.write('K7','IDH1hotspotR132',plaincellformat)
    worksheet.write('K8','KIThotspotD816V',plaincellformat)
    worksheet.write('K9','NPM1hotspot',plaincellformat)
    worksheet.write('K10','JAK2hotspotexon12',plaincellformat)
    worksheet.write('K11','JAK2hotspotV617F',plaincellformat)
    worksheet.write('K12','HRAShotspot61',plaincellformat)
    worksheet.write('K13','HRAShotspot12&13',plaincellformat)
    worksheet.write('K14','CBLhotspotC381R420',plaincellformat)
    worksheet.write('K15','KRAShotspot61',plaincellformat)
    worksheet.write('K16','KRAShotspot12&13',plaincellformat)
    worksheet.write('K17','IDH2hotspotR172',plaincellformat)
    worksheet.write('K18','SRSF2hotspotP95',plaincellformat)
    worksheet.write('K19','SETBP1hotspot868880',plaincellformat)
    worksheet.write('K20','U2AF1hotspotQ157',plaincellformat)
    worksheet.write('K21','U2AF1hotspotS34',plaincellformat)
    worksheet.write('K22','IDH2hotspotR140',plaincellformat)

    worksheet.write('K1','Hotspot coverage (100x)',kcolumnheaderformat)

    #Write coverage data to Excel spreadsheet
    fillcell(worksheet,coveragedata,plaincellformat,'B2','BRAF')
    fillcell(worksheet,coveragedata,plaincellformat,'B3','CALR')
    fillcell(worksheet,coveragedata,plaincellformat,'B4','CBL')
    fillcell(worksheet,coveragedata,plaincellformat,'B5','CBLB')
    fillcell(worksheet,coveragedata,plaincellformat,'B6','CSF3R')
    fillcell(worksheet,coveragedata,plaincellformat,'B7','FLT3')
    fillcell(worksheet,coveragedata,plaincellformat,'B8','HRAS')
    fillcell(worksheet,coveragedata,plaincellformat,'B9','JAK2')
    fillcell(worksheet,coveragedata,plaincellformat,'B10','JAK3')
    fillcell(worksheet,coveragedata,plaincellformat,'B11','KIT')
    fillcell(worksheet,coveragedata,plaincellformat,'B12','KRAS')
    fillcell(worksheet,coveragedata,plaincellformat,'B13','MPL')
    fillcell(worksheet,coveragedata,plaincellformat,'B14','NRAS')
    fillcell(worksheet,coveragedata,plaincellformat,'B15','PDGFRA')
    fillcell(worksheet,coveragedata,plaincellformat,'B16','PTPN11')

    #These remove GNAS
    fillcell(worksheet,coveragedata,plaincellformat,'D2','MYD88')
    fillcell(worksheet,coveragedata,plaincellformat,'D3','NOTCH1')
    fillcell(worksheet,coveragedata,plaincellformat,'D4','PTEN')
    fillcell(worksheet,coveragedata,plaincellformat,'D7','ASXL1')
    fillcell(worksheet,coveragedata,plaincellformat,'D8','ATRX')
    fillcell(worksheet,coveragedata,plaincellformat,'D9','BCOR')
    fillcell(worksheet,coveragedata,plaincellformat,'D10','BCORL1')
    fillcell(worksheet,coveragedata,plaincellformat,'D11','DNMT3A')
    fillcell(worksheet,coveragedata,plaincellformat,'D12','EZH2')
    fillcell(worksheet,coveragedata,plaincellformat,'D13','GATA1')
    fillcell(worksheet,coveragedata,plaincellformat,'D14','IDH1')
    fillcell(worksheet,coveragedata,plaincellformat,'D15','IDH2')
    fillcell(worksheet,coveragedata,plaincellformat,'D16','KMT2A')
    fillcell(worksheet,coveragedata,plaincellformat,'D17','PHF6')
    fillcell(worksheet,coveragedata,plaincellformat,'D18','TET2')
    
    #These are old versions prior to removing GNAS
    #fillcell(worksheet,coveragedata,plaincellformat,'D2','GNAS')
    #fillcell(worksheet,coveragedata,plaincellformat,'D3','MYD88')
    #fillcell(worksheet,coveragedata,plaincellformat,'D4','NOTCH1')
    #fillcell(worksheet,coveragedata,plaincellformat,'D5','PTEN')
    #fillcell(worksheet,coveragedata,plaincellformat,'D8','ASXL1')
    #fillcell(worksheet,coveragedata,plaincellformat,'D9','ATRX')
    #fillcell(worksheet,coveragedata,plaincellformat,'D10','BCOR')
    #fillcell(worksheet,coveragedata,plaincellformat,'D11','BCORL1')
    #fillcell(worksheet,coveragedata,plaincellformat,'D12','DNMT3A')
    #fillcell(worksheet,coveragedata,plaincellformat,'D13','EZH2')
    #fillcell(worksheet,coveragedata,plaincellformat,'D14','GATA1')
    #fillcell(worksheet,coveragedata,plaincellformat,'D15','IDH1')
    #fillcell(worksheet,coveragedata,plaincellformat,'D16','IDH2')
    #fillcell(worksheet,coveragedata,plaincellformat,'D17','KMT2A')
    #fillcell(worksheet,coveragedata,plaincellformat,'D18','PHF6')
    #fillcell(worksheet,coveragedata,plaincellformat,'D19','TET2')
    fillcell(worksheet,coveragedata,plaincellformat,'F2','RAD21')
    fillcell(worksheet,coveragedata,plaincellformat,'F3','SMC1A')
    fillcell(worksheet,coveragedata,plaincellformat,'F4','SMC3')
    fillcell(worksheet,coveragedata,plaincellformat,'F5','STAG2')
    fillcell(worksheet,coveragedata,plaincellformat,'F8','CUX1')
    fillcell(worksheet,coveragedata,plaincellformat,'F9','ETV6')
    fillcell(worksheet,coveragedata,plaincellformat,'F10','GATA2')
    fillcell(worksheet,coveragedata,plaincellformat,'F11','IKZF1')
    fillcell(worksheet,coveragedata,plaincellformat,'F12','RUNX1')
    fillcell(worksheet,coveragedata,plaincellformat,'F13','WT1')
    fillcell(worksheet,coveragedata,plaincellformat,'H2','SF3B1')
    fillcell(worksheet,coveragedata,plaincellformat,'H3','SRSF2')
    fillcell(worksheet,coveragedata,plaincellformat,'H4','U2AF1')
    fillcell(worksheet,coveragedata,plaincellformat,'H5','ZRSR2')
    fillcell(worksheet,coveragedata,plaincellformat,'H8','FBXW7')
    fillcell(worksheet,coveragedata,plaincellformat,'H9','NPM1')
    fillcell(worksheet,coveragedata,plaincellformat,'H10','SETBP1')
    fillcell(worksheet,coveragedata,plaincellformat,'H11','TP53')
    fillcell(worksheet,coveragedata,plaincellformat,'L2','NRAShotspot12&13')
    fillcell(worksheet,coveragedata,plaincellformat,'L3','DNMT3AhotspotR882')
    fillcell(worksheet,coveragedata,plaincellformat,'L4','SF3B1hotspotK700')
    fillcell(worksheet,coveragedata,plaincellformat,'L5','SF3B1hotspotH662K666')
    fillcell(worksheet,coveragedata,plaincellformat,'L6','SF3B1hotspotE622N626')
    fillcell(worksheet,coveragedata,plaincellformat,'L7','IDH1hotspotR132')
    fillcell(worksheet,coveragedata,plaincellformat,'L8','KIThotspotD816V')
    fillcell(worksheet,coveragedata,plaincellformat,'L9','NPM1hotspot')
    fillcell(worksheet,coveragedata,plaincellformat,'L10','JAK2hotspotexon12')
    fillcell(worksheet,coveragedata,plaincellformat,'L11','JAK2hotspotV617F')
    fillcell(worksheet,coveragedata,plaincellformat,'L12','HRAShotspot61')
    fillcell(worksheet,coveragedata,plaincellformat,'L13','HRAShotspot12&13')
    fillcell(worksheet,coveragedata,plaincellformat,'L14','CBLhotspotC381R420')
    fillcell(worksheet,coveragedata,plaincellformat,'L15','KRAShotspot61')
    fillcell(worksheet,coveragedata,plaincellformat,'L16','KRAShotspot12&13')
    fillcell(worksheet,coveragedata,plaincellformat,'L17','IDH2hotspotR172')
    fillcell(worksheet,coveragedata,plaincellformat,'L18','SRSF2hotspotP95')
    fillcell(worksheet,coveragedata,plaincellformat,'L19','SETBP1hotspot868880')
    fillcell(worksheet,coveragedata,plaincellformat,'L20','U2AF1hotspotQ157')
    fillcell(worksheet,coveragedata,plaincellformat,'L21','U2AF1hotspotS34')
    fillcell(worksheet,coveragedata,plaincellformat,'L22','IDH2hotspotR140')
