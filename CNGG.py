import xlsxwriter
from pathlib import Path
import json
import tkinter as tk
from tkinter import ttk
from datetime import datetime
griddy = json.load(Path("griddy.json").open("r", encoding="UTF-8")) if Path("griddy.json").exists() == True else []

def cngg(inp:int = 100, ind:bool = False) -> list:
    """
    Args:
        inp (int, optional): Sets the maximum number for the grid. Defaults to 100.
        ind (bool, optional): If True, a grid will be exported to a separate Excel file, with the input number noted in brackets in the file name. Defaults to False.

    Returns:
        list: List containing the numbers in the grid from 1 to the number specified in `inp`, dictionary of square numbers with the squared number as the key, and the number squared as their value, dictionary of griddy numbers (numbers which produce a clean grid allowing for a clean cross grid with no additional rows), and the number used to set the total number of rows and columns
    """
    inp = int(inp)
    numList = list(range(1,inp + 1))
    squares = {} # this dictionary stores the result of squaring numbers up to the value of `inp` in the key, whose value is the number squared, which is then searched to identify a squared number closer to `inp`, whose value is then set as the total columns and rows to have in the grid, to ensure griddiness as much as possible.
    squaresInGrid = {} # this stores the row and column information for square numbers in the grid which then populates a reference table below the grid
    sq = 1
    savePath = ""
    while sq <= inp:
        squares[sq * sq] = sq
        if (sq * sq) > 0 and (inp == sq * sq or (inp - (sq * sq) < 10 and (sq + 1) * (sq + 1) != inp)):
            rowsCol = sq
            sq = inp + 1 # this allows the while loop to terminate by maxing the value of `sq`
        else:
            sq += 1

    with xlsxwriter.Workbook(f"{"CNGG" if ind == False else "CNGG (" + str(inp) + ")"}.xlsx") as xl:
        ws = xl.add_worksheet("Grid")
        savePath = f"{Path(__file__).parent.as_posix()}/{xl.filename}"
        row = 0
        col = 0
        LtoRCrossGridSum = 0
        RtoLCrossGridSum = 0
        MidRowSum = 0
        MidColSum = 0
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
                
                if col == rowsCol and row > 1 and eachNum > 9 and griddy.count(eachNum) == None: # 9 does not produce a grid, only a row, but is matched for some reason so manually excluded
                    griddy.append(eachNum)

                # No elifs below to record crossgrid numbers that can be considered both left to right and right to left
                if eachNumPos == (rowsCol * row) - row:
                    RtoLCrossGridSum += eachNum
                
                if row == col:
                    LtoRCrossGridSum += eachNum

                if isMidRow:
                    MidRowSum += eachNum

                if isMidCol:
                    MidColSum += eachNum
            elif squares.get(eachNum) != None:
                ws.write_row(row, col, [eachNum], xl.add_format({"align": "center", "bold": True, "font_color": "red", "font_size": 17, "font_name": "Cascadia Code"}))

                if isMidRow:
                    MidRowSum += eachNum

                if isMidCol:
                    MidColSum += eachNum

                ws.data_validation(row, col, row, col, {"validate": "integer", "criteria": ">", "value": -1, "input_title": f"{eachNum} (square)" if len(str(eachNum)) < 24 else "Square","input_message": f"Row {row}, Column {col}"})
                squaresInGrid[eachNum] = [row, col]
            else:
                if isMidCol:
                    ws.write_row(row, col, [eachNum], xl.add_format({"align": "center", "bold": True, "font_color": "green", "font_size": 17, "font_name": "Cascadia Code"}))
                    
                    ws.data_validation(row, col, row, col, {"validate": "integer", "criteria": ">", "value": -1, "input_title": f"{eachNum} (midcol)" if len(str(eachNum)) < 24 else "Mid Col","input_message": f"Row {row}, Column {col}"})

                    MidColSum += eachNum
                elif isMidRow:
                    ws.write_row(row, col, [eachNum], xl.add_format({"align": "center", "bold": True, "font_color": "green", "font_size": 17, "font_name": "Cascadia Code"}))
                    
                    ws.data_validation(row, col, row, col, {"validate": "integer", "criteria": ">", "value": -1, "input_title": f"{eachNum} (midrow)" if len(str(eachNum)) < 24 else "Mid Row","input_message": f"Row {row}, Column {col}"})

                    MidRowSum += eachNum
                else:
                    ws.write_row(row, col, [eachNum], xl.add_format({"align": "center", "font_color": "#E4DFEC", "font_name": "Cascadia Code"}))
                    
                    ws.data_validation(row, col, row, col, {"validate": "integer", "criteria": ">", "value": -1, "input_title": f"{eachNum}" if len(str(eachNum)) < 33 else "","input_message": f"Row {row}, Column {col}"})
            col += 1
        ws.freeze_panes(1,1)
        ws.hide_gridlines(2)

        ws.write(row + 2,1, f"Numbers between 1 and {inp}, with square numbers ({len(squaresInGrid)} in grid) coloured in red, crossgrid numbers coloured in orange, and middle column and middle row numbers coloured in green, unless they are square numbers.")

        ws.merge_range(row + 5, 7, row + 5, 9, "Left to right crossgrid sum", xl.add_format({"bold": True, "align": "center"}))
        ws.write_row(row + 5, 10, [LtoRCrossGridSum], xl.add_format({"align": "center"}))
        ws.merge_range(row + 6, 7, row + 6, 9, "Right to left crossgrid sum", xl.add_format({"bold": True, "align": "center"}))
        ws.write_row(row + 6, 10, [RtoLCrossGridSum], xl.add_format({"align": "center"}))
        ws.merge_range(row + 7, 7, row + 7, 9, "Middle row sum", xl.add_format({"bold": True, "align": "center"}))
        ws.write_row(row + 7, 10, [MidRowSum], xl.add_format({"align": "center"}))
        ws.merge_range(row + 8, 7, row + 8, 9, "Middle column sum", xl.add_format({"bold": True, "align": "center"}))
        ws.write_row(row + 8, 10, [MidColSum], xl.add_format({"align": "center"}))
        if griddy.count(inp) > 0:
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

    return [numList, squares, griddy, rowsCol, savePath]

def gui():
    cnggInterface = tk.Tk()
    cnggInterface.resizable(0,0)
    ttkStyle = ttk.Style()
    allThemes = ttkStyle.theme_names()
    currentTheme = tk.StringVar(value = "default")
    cnggInterface.title("Colourful Number Grid Generator")
    inputFieldNum = tk.StringVar(value="100")
    cnggInpLabel = ttk.Label(cnggInterface, text = "Input number")
    cnggInp = ttk.Entry(cnggInterface, textvariable=inputFieldNum)
    saveAsIndFilesBool = tk.BooleanVar(value = False)
    
    def saveAsCheckBoxOnClick():
        print(f"{datetime.now().time().strftime("%H:%M:%S")}: Individual export to Excel set to: {saveAsIndFilesBool.get()}")

    saveAsCheckbox = ttk.Checkbutton(cnggInterface, variable = saveAsIndFilesBool, text = "Export to separate Excel", command = saveAsCheckBoxOnClick)

    def interfaceSubmit():
        run = cngg(cnggInp.get(), saveAsIndFilesBool.get())
        print(f"{datetime.now().time().strftime("%H:%M:%S")}: Saved CNGG export at {run[-1]}")

    submitButton = ttk.Button(cnggInterface, text = "Run CNGG", command = interfaceSubmit)

    cnggInpLabel.grid(row = 1, column = 0)
    cnggInp.grid(row = 1, column = 1)
    saveAsCheckbox.grid(row = 1, column = 2)
    submitButton.grid(row = 1, column = 3)

    def selectStyle(selectedStyle):
        ttkStyle.theme_use(selectedStyle)
        print(f"{datetime.now().time().strftime("%H:%M:%S")}: Theme set to: {selectedStyle}")

    styleDropdownLabel = ttk.Label(cnggInterface, text = "Select style")
    styleDropdown = ttk.OptionMenu(cnggInterface, currentTheme, *allThemes, command = selectStyle)
    styleDropdownLabel.grid(row = 2, column = (2 if len(griddy) > 0 else 0))
    styleDropdown.grid(row = 2, column = (3 if len(griddy) > 0 else 1))

    if len(griddy) > 0:
        griddyNumber = tk.StringVar(value = griddy[0] if len(griddy) > 0 else 0)

        def selectGriddy(selectedGriddyNumber):
            inputFieldNum.set(selectedGriddyNumber)
            print(f"{datetime.now().time().strftime("%H:%M:%S")}: Griddy number selected: {selectedGriddyNumber}")

        griddyDropdownLabel = ttk.Label(cnggInterface, text = "Select griddy number")
        griddyDropdown = ttk.OptionMenu(cnggInterface, griddyNumber, *griddy, command=selectGriddy)
        griddyDropdownLabel.grid(row = 2, column = 0)
        griddyDropdown.grid(row = 2, column = 1)



    cnggInterface.mainloop()