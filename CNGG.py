import xlsxwriter
from pathlib import Path
import json
import tkinter as tk
griddy = json.load(Path("griddy.json").open("r", encoding="UTF-8")) if Path("griddy.json").exists() == True else {}
def cngg(inp:int = 100, ind:bool = False):
    inp = int(inp)
    numList = list(range(1,inp + 1))
    squares = {} # this dictionary stores the result of squaring numbers up to the value of `inp` in the key, whose value is the number squared, which is then searched to identify a squared number closer to `inp`, whose value is then set as the total columns and rows to have in the grid, to ensure griddiness as much as possible.
    squaresInGrid = {} # this stores the row and column information for square numbers in the grid which then populates a reference table below the grid
    sq = 1
    while sq <= inp:
        squares[sq * sq] = sq
        if (sq * sq) > 0 and (inp == sq * sq or inp - (sq * sq) < 10):
            rowsCol = sq
            sq = inp + 1 # this allows the while loop to terminate by maxing the value of `sq`
        else:
            sq += 1

    with xlsxwriter.Workbook(f"{"CNGG" if ind == False else "CNGG (" + str(inp) + ")"}.xlsx") as xl:
        ws = xl.add_worksheet("Grid")
        row = 0
        col = 0
        for eachNumPos, eachNum in enumerate(numList):
            if row == 0:
                ws.write_row(row, 1, [f"Col {colNum}" for colNum in list(range(1, rowsCol + 1))], xl.add_format({"align": "center", "font_color": "#E4DFEC", "font_name": "Cascadia Code", "italic": True, "bottom": True}))
            if eachNumPos % rowsCol == 0: # the first eachNumPos of 0 will match to 0 meaning the very first row of the grid will start on row 1. Subsequent number positions are evaluated to identify where they are divisible fully by `rowsCol` so the next set of numbers are populated in the next row, starting from column 1 (B).
                row += 1
                col = 1
            if row > 0: # this populates the row number header row
                ws.write_row(row, 0, [f"Row {row}"], xl.add_format({"align": "center", "font_color": "#E4DFEC", "font_name": "Cascadia Code", "italic": True, "right": True}))
            
            isMidCol = (col == int((rowsCol + 1) / 2) and rowsCol % 2 == 1) or (rowsCol % 2 == 0 and (col == int(rowsCol / 2) or col == int(rowsCol / 2) + 1)) # for odd numbered columns, check if the current column is half of the total rows + 1 (which converts the total rows to an even number). For odd numbered rows, it is easier to cleanly mark the center column as there is an equal number of columns to the left and right of it. For even numbered columns, check if the current column is the half of total rows or the column next to the half column to highlight both which then equalises the number of columns to their left and right.

            isMidRow = (row == int((rowsCol + 1) / 2) and rowsCol % 2 == 1) or (rowsCol % 2 == 0 and (row == int(rowsCol / 2) or row == int(rowsCol / 2) + 1)) # similar to isMidRow but just for rows
            
            if eachNumPos == ((rowsCol * row) - row) or row == col: 
                ws.write_row(row, col, [eachNum], xl.add_format({"align": "center", "bold": True, "font_color": "orange", "font_size": 17, "font_name": "Cascadia Code"})) if isMidRow == False else ws.write_row(row, col, [eachNum], xl.add_format({"align": "center", "bold": True, "bg_color": "#002060", "font_color": "white", "font_size": 17, "font_name": "Cascadia Code", "diag_color": "red"})) # the else condition highlights the central number in the grid
                
                ws.data_validation(row, col, row, col, {"validate": "integer", "criteria": ">", "value": -1, "input_title": f"{eachNum} (crossgrid)" if len(str(eachNum)) < 24 else "Cross Grid","input_message": f"Row {row}, Column {col}"}) if isMidRow == False else ws.data_validation(row, col, row, col, {"validate": "integer", "criteria": ">", "value": -1, "input_title": f"{eachNum} (middle value)" if len(str(eachNum)) < 18 else "Cross Grid Middle Value","input_message": f"Row {row}, Column {col}"})
                
                if col == rowsCol and row > 1 and eachNum > 9 and griddy.get(str(eachNum)) == None: # 9 does not produce a grid, only a row, but is matched for some reason so manually excluded
                    griddy[eachNum] = eachNum if len(griddy) == 0 else eachNum - int(list(griddy.keys())[-1])
            elif squares.get(eachNum) != None:
                ws.write_row(row, col, [eachNum], xl.add_format({"align": "center", "bold": True, "font_color": "red", "font_size": 17, "font_name": "Cascadia Code"}))
                
                ws.data_validation(row, col, row, col, {"validate": "integer", "criteria": ">", "value": -1, "input_title": f"{eachNum} (square)" if len(str(eachNum)) < 24 else "Square","input_message": f"Row {row}, Column {col}"})
                squaresInGrid[eachNum] = [row, col]
            else:
                if isMidCol:
                    ws.write_row(row, col, [eachNum], xl.add_format({"align": "center", "bold": True, "font_color": "green", "font_size": 17, "font_name": "Cascadia Code"}))
                    
                    ws.data_validation(row, col, row, col, {"validate": "integer", "criteria": ">", "value": -1, "input_title": f"{eachNum} (midcol)" if len(str(eachNum)) < 24 else "Mid Col","input_message": f"Row {row}, Column {col}"})
                elif isMidRow:
                    ws.write_row(row, col, [eachNum], xl.add_format({"align": "center", "bold": True, "font_color": "green", "font_size": 17, "font_name": "Cascadia Code"}))
                    
                    ws.data_validation(row, col, row, col, {"validate": "integer", "criteria": ">", "value": -1, "input_title": f"{eachNum} (midrow)" if len(str(eachNum)) < 24 else "Mid Row","input_message": f"Row {row}, Column {col}"})
                else:
                    ws.write_row(row, col, [eachNum], xl.add_format({"align": "center", "font_color": "#E4DFEC", "font_name": "Cascadia Code"}))
                    
                    ws.data_validation(row, col, row, col, {"validate": "integer", "criteria": ">", "value": -1, "input_title": f"{eachNum}" if len(str(eachNum)) < 33 else "","input_message": f"Row {row}, Column {col}"})
            col += 1
        ws.freeze_panes(1,1)
        ws.hide_gridlines(2)

        ws.write(row + 2,1, f"Numbers between 1 and {inp}, with square numbers ({len(squaresInGrid)} in grid) coloured in red, cross-grid numbers coloured in orange, and middle column and middle row numbers coloured in green, unless they are square numbers.")
        if griddy.get(str(inp)) != None:
            ws.write_row(0,0,["GRIDDY"], xl.add_format({"bold": True, "font_name": "Cascadia Code", "font_color": "red", "align": "center"}))
        ws.write_row(row + 4, 1, ["Table of square numbers in grid and location"], xl.add_format({"bold": True, "font_size": 12, "font_name": "Cascadia Code"}))
        i = 6
        for squ in squaresInGrid:
            ws.write_row(row + i, 1, [int(squ), squaresInGrid[squ][0], squaresInGrid[squ][1]], xl.add_format({"font_name": "Cascadia Code", "align": "center"}))
            i += 1
        ws.add_table(row + 5, 1, row + i, 3, {"total_row": True, "columns": [{"header": col, "total_function": "count"} if col == "#" else {"header": col} for col in ["#", "Row", "Col"]], "style": "Table Style Light 8"})
        ws.set_row(row + 5,cell_format=xl.add_format({"align": "center"}))
        
    if len(griddy) > 0:
        with Path("griddy.json").open("w", encoding="UTF-8") as gridJSON:
            gridJSON.write(json.dumps(griddy, indent=2, ensure_ascii=False))

    return [numList, squares, griddy, rowsCol]

def gui():
    cnggInterface = tk.Tk()
    cnggInterface.title("Colourful Number Grid Generator")
    inputFieldNum = tk.StringVar(value="100")
    cnggInpLabel = tk.Label(cnggInterface, text = "Input number")
    cnggInp = tk.Entry(cnggInterface, textvariable=inputFieldNum)
    saveAsIndFiles = tk.BooleanVar(value = False)
    saveAsCheckbox = tk.Checkbutton(cnggInterface, variable = saveAsIndFiles, text = "Export to separate Excel")

    def interfaceSubmit():
        cngg(cnggInp.get(), saveAsIndFiles.get())

    submit = tk.Button(cnggInterface, text = 'Run CNGG', command = interfaceSubmit)
    cnggInpLabel.grid(row = 1, column = 0)
    cnggInp.grid(row = 1, column = 1)
    saveAsCheckbox.grid(row = 1, column = 2)
    submit.grid(row = 1, column = 3)

    def showGriddyNumbers():
        def selectGriddy(event):
            inputFieldNum.set(list(griddy.keys())[reses.curselection()[0] - 1])
        griddyWindow = tk.Tk()
        griddyWindow.title("List of griddy inputs")
        scrl = tk.Scrollbar(griddyWindow)
        scrl.pack(side='right', fill='y')
        reses = tk.Listbox(griddyWindow,yscrollcommand=scrl.set)
        reses.bind("<Double-1>", selectGriddy)
        reses.insert(0, "Numbers")
        griddyWindow.minsize(200,100)
        colNum = 0
        for x in list(griddy.keys()):
            reses.insert(colNum+1, f"{colNum}. {x}")
            colNum += 1
        reses.pack(side="left", fill = "both")
        scrl.config(command = reses.yview)
        
    seeGriddy = tk.Button(cnggInterface, text = 'See griddy numbers', command = showGriddyNumbers)
    seeGriddy.grid(row = 2, column = 0)
    cnggInterface.mainloop()