
def adjust_column_dimensions(worksheet):
    """
Receives a worksheet and adjusts the width of each column based on the widest data in the column, plus 2.
    """    
    for column in worksheet.columns: 
        
        max_width = 0
        col_letter = column[0].column_letter
        
        
        for line in column:
            
            value = str(line.value)
            current_width = len(value)
                
            if current_width > max_width:
                max_width = current_width
        
        worksheet.column_dimensions[col_letter].width = max_width + 4
            