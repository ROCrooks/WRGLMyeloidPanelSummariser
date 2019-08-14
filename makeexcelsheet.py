import xlsxwriter

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
    worksheet.merge_range('C6:D7','DNA/chromatin\nmodification',mergedheadersformat)
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
    worksheet.write('C2','GNAS',italicgenelabelcellformat)
    worksheet.write('C3','MYD88',italicgenelabelcellformat)
    worksheet.write('C4','NOTCH1',italicgenelabelcellformat)
    worksheet.write('C5','PTEN',italicgenelabelcellformat)
    worksheet.write('C8','ASXL1',italicgenelabelcellformat)
    worksheet.write('C9','ATRX',italicgenelabelcellformat)
    worksheet.write('C10','BCOR',bolditalicgenelabelcellformat)
    worksheet.write('C11','BCORL1',bolditalicgenelabelcellformat)
    worksheet.write('C12','DNMT3A',bolditalicgenelabelcellformat)
    worksheet.write('C13','EZH2',bolditalicgenelabelcellformat)
    worksheet.write('C14','GATA1',italicgenelabelcellformat)
    worksheet.write('C15','IDH1',italicgenelabelcellformat)
    worksheet.write('C16','IDH2',italicgenelabelcellformat)
    worksheet.write('C17','KMT2A',italicgenelabelcellformat)
    worksheet.write('C18','PHF6',bolditalicgenelabelcellformat)
    worksheet.write('C19','TET2',italicgenelabelcellformat)
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
