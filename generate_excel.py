import xlsxwriter
import json

def load_config():
    with open('config.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def create_excel():
    config = load_config()
    players = config['players']
    matches = config['matches']
    
    workbook = xlsxwriter.Workbook('predictions.xlsx')
    worksheet = workbook.add_worksheet('Predictions')
    
    # Formats
    bold = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
    header = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#D7E4BC', 'border': 1})
    match_header = workbook.add_format({'bold': True, 'align': 'left', 'bg_color': '#FFEB9C', 'border': 1, 'font_size': 14})
    cell = workbook.add_format({'align': 'center', 'border': 1})
    
    worksheet.set_column('A:A', 20)
    worksheet.set_column('B:F', 15)
    
    row = 1
    
    for match in matches:
        home, away = match['home'], match['away']
        
        # Match header
        worksheet.merge_range(row, 0, row, 2, f"{home} vs {away}", match_header)
        worksheet.write(row, 3, "Real Score:", bold)
        
        real_home = xlsxwriter.utility.xl_rowcol_to_cell(row, 4)
        real_away = xlsxwriter.utility.xl_rowcol_to_cell(row, 5)
        
        worksheet.write(row, 4, "", cell)
        worksheet.write(row, 5, "", cell)
        
        row += 2
        
        # Table headers
        for col, h in enumerate(['Name', 'Pred Home', 'Pred Away', 'Points']):
            worksheet.write(row, col, h, header)
        row += 1
        
        # Player rows
        for player in players:
            pred = match['predictions'].get(player, [0, 0])
            
            worksheet.write(row, 0, player, cell)
            worksheet.write(row, 1, pred[0], cell)
            worksheet.write(row, 2, pred[1], cell)
            
            # Points formula
            ph = xlsxwriter.utility.xl_rowcol_to_cell(row, 1)
            pa = xlsxwriter.utility.xl_rowcol_to_cell(row, 2)
            
            exact = f"AND({real_home}={ph},{real_away}={pa})"
            home_win = f"AND({real_home}>{real_away},{ph}>{pa})"
            away_win = f"AND({real_home}<{real_away},{ph}<{pa})"
            draw = f"AND({real_home}={real_away},{ph}={pa})"
            
            formula = f'=IF(OR(ISBLANK({real_home}),ISBLANK({real_away})),"",IF({exact},3,IF(OR({home_win},{away_win},{draw}),1,0)))'
            worksheet.write_formula(row, 3, formula, cell)
            
            row += 1
        
        row += 2

    # Total points section
    row += 1
    total_fmt = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1, 'font_size': 14})
    total_cell = workbook.add_format({'bold': True, 'align': 'center', 'border': 1, 'bg_color': '#D9E2F3'})
    
    worksheet.merge_range(row, 0, row, 3, "TOTAL POINTS", total_fmt)
    row += 2
    
    worksheet.write(row, 0, "Name", header)
    worksheet.write(row, 1, "Points", header)
    row += 1
    
    for player in players:
        worksheet.write(row, 0, player, cell)
        worksheet.write_formula(row, 1, f'=SUMIF(A:A,"{player}",D:D)', total_cell)
        row += 1

    workbook.close()
    print("Excel file 'predictions.xlsx' created successfully!")

if __name__ == "__main__":
    create_excel()
