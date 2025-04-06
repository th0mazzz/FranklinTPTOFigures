import os
import math
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

# DEFINING FUNCTIONS
def colLettersToNumber(letters):
    letter_list = list('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
    if letters in letter_list:
        return letter_list.index(letters) + 1
    else:
        running_total = 0
        place = len(letters) - 1
        for l in letters:
            running_total = running_total + (26 ** place) * (letter_list.index(l) + 1)
            place = place - 1

        return running_total

def colNumberToLetters(n):
    column_letter = ''
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        column_letter = chr(65 + remainder) + column_letter

    return column_letter

def relativeToCell(cell_string, translations):
    # Returns the cell per the list of translations
    # e.g., relativeToCell(ws, 'B2', [('right', 4), ('down', 2)]) should return 'F4'

    current_letters, current_numbers = splitCellCoord(cell_string)
    current_numbers = int(current_numbers)

    for movement in translations:
        dir = movement[0]
        units = movement[1]
        if dir == 'right':
            current_letters = colNumberToLetters(colLettersToNumber(current_letters) + units)
        elif dir == 'left':
            current_letters = colNumberToLetters(colLettersToNumber(current_letters) - units)
        elif dir == 'down':
            current_numbers = current_numbers + units
        elif dir == 'up':
            current_numbers = current_numbers - units

    if current_letters == '':
        raise Exception('Target column is out of range')   
    if current_numbers < 0:
        raise Exception("Target row is out of range")
    
    return current_letters + str(current_numbers)

def splitCellCoord(coord):
    i_cell = len(coord) - 1
    while coord[i_cell].isnumeric():
        i_cell = i_cell - 1

    cell_letters = coord[:i_cell + 1]
    cell_numbers = coord[i_cell + 1:]

    return (cell_letters, cell_numbers)

def fillCellColors(worksheet, range_start, range_end, color, f_type):
    # Fills cell colors in 'worksheet' within the specified rectangular 
    # range [range_start:range_end] with color 'color' and fill type 'f_type' 

    fill = PatternFill(start_color = color, end_color = color, fill_type = f_type)
    for row in worksheet[range_start + ":" + range_end]:
        for cell in row:
            cell.fill = fill

def setColumnWidths(worksheet, desired_width):
    # Sets widths of columns A through Z in 'worksheet' to be 'desired_width'

    cols = list('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
    for col in cols:
        worksheet.column_dimensions[col].width = desired_width

def createThickOutsideBorders(worksheet, range_start, range_end):
    # Creates thick outside borders on rectnangular range [range_start, range_end]
    # in 'worksheet' THIS ALGO CAN BE MADE MORE EFFICIENT WITH THE RELATIVETOCELL FUNCTION I MADE.
    
    start_letters, start_numbers = splitCellCoord(range_start)
    end_letters, end_numbers = splitCellCoord(range_end)

    for row in worksheet[range_start + ":" + range_end]:
        for cell in row: 

            coord = cell.coordinate

            cell_letters, cell_numbers = splitCellCoord(coord)

            # corners 
            if cell_letters == start_letters and cell_numbers == start_numbers:
                worksheet[coord].border = Border(left=Side(style='thick'), top=Side(style='thick'))
            elif cell_letters == start_letters and cell_numbers == end_numbers:
                worksheet[coord].border = Border(left=Side(style='thick'), bottom=Side(style='thick'))
            elif cell_letters == end_letters and cell_numbers == start_numbers:
                worksheet[coord].border = Border(right=Side(style='thick'), top=Side(style='thick'))
            elif cell_letters == end_letters and cell_numbers == end_numbers:
                worksheet[coord].border = Border(right=Side(style='thick'), bottom=Side(style='thick'))
            # edges
            elif cell_letters == start_letters:
                worksheet[coord].border = Border(left=Side(style='thick'))
            elif cell_letters == end_letters:
                worksheet[coord].border = Border(right=Side(style='thick'))
            elif cell_numbers == start_numbers:
                worksheet[coord].border = Border(top=Side(style='thick'))
            elif cell_numbers == end_numbers:
                worksheet[coord].border = Border(bottom=Side(style='thick'))
            
def createIntersectionBorders(worksheet, main_coords):
    # Creates intersection borders denoting the roadways in 'worksheet'

    top_left_letter, top_left_number = splitCellCoord(main_coords[0])
    bottom_right_letter, bottom_right_number = splitCellCoord(main_coords[1])

    top_left_number = int(top_left_number)
    bottom_right_number = int(bottom_right_number)

    if main_display_width % 2 == 0:
        hori_midpoint = math.floor((colLettersToNumber(bottom_right_letter) - colLettersToNumber(top_left_letter))/2 + colLettersToNumber(top_left_letter))

    else: 
        hori_midpoint = math.floor((colLettersToNumber(bottom_right_letter) - colLettersToNumber(top_left_letter))/2 + colLettersToNumber(top_left_letter))
    if main_display_height % 2 == 0:
        vert_midpoint = math.floor((bottom_right_number - top_left_number)/2 + top_left_number)
    else: 
        vert_midpoint = math.floor((colLettersToNumber(bottom_right_letter) - colLettersToNumber(top_left_letter))/2 + colLettersToNumber(top_left_letter))

    hori_midpoint_cell = colNumberToLetters(hori_midpoint) + str(top_left_number + 1)
    vert_midpoint_cell = colNumberToLetters(colLettersToNumber(top_left_letter) + 1) + str(vert_midpoint)
    
    vert_end = relativeToCell(hori_midpoint_cell, [('down', main_display_height-3)])
    hori_end = relativeToCell(vert_midpoint_cell, [('right', main_display_width-3)])

    current_vert = hori_midpoint_cell
    reachedVertEnd = False
    while not reachedVertEnd:
        if current_vert == vert_end:
            reachedVertEnd = True
        worksheet[current_vert].border = Border(right=Side(style='thin'))
        current_vert = relativeToCell(current_vert, [('down', 1)])

    current_hori = vert_midpoint_cell
    reachedHoriEnd = False
    while not reachedHoriEnd:
        if current_hori == hori_end:
            reachedHoriEnd = True
        worksheet[current_hori].border = Border(bottom=Side(style='thin'))
        current_hori = relativeToCell(current_hori, [('right', 1)])

    middle_cell = splitCellCoord(hori_midpoint_cell)[0] + str(splitCellCoord(vert_midpoint_cell)[1])
    ws[middle_cell].border = Border(bottom=Side(style='thin'), right=Side(style='thin'))    

def generateFigure(ws, origin):

    # CREATE LAYOUT IN EXCEL SPREADSHEET
    header_coords = [None, None]
    main_coords = [None, None]

    if header:
        # header_height = 2
        header_coords = [origin, relativeToCell(origin, [('right', main_display_width-1), ('down', header_height-1)])]
        main_coords = [relativeToCell(origin, [('down', 2)]), relativeToCell(origin, [('right', main_display_width-1), ('down', main_display_height + header_height-1)])]
        fillCellColors(ws, header_coords[0], header_coords[1], header_color, 'solid')
        createThickOutsideBorders(ws, header_coords[0], header_coords[1])
    else:
        main_coords = [origin, relativeToCell(origin, [('right', main_display_width-1), ('down', main_display_height-1)])]

    if main_border:
        fillCellColors(ws, main_coords[0], main_coords[1], main_border_color, 'solid')
        fillCellColors(ws, relativeToCell(main_coords[0], [('right', 1), ('down', 1)]), relativeToCell(main_coords[1], [('left', 1), ('up', 1)]), main_background_color, 'solid')
    else:
        fillCellColors(ws, main_coords[0], main_coords[1], main_background_color, 'solid')

    # print('origin: ' + str(origin))
    # print('header: ' + str(header_coords))
    # print('main: ' + str(main_coords))


    if cardinal_dirs:
        top_left_letter, top_left_number = splitCellCoord(main_coords[0])
        bottom_right_letter, bottom_right_number = splitCellCoord(main_coords[1])

        top_left_number = int(top_left_number)
        bottom_right_number = int(bottom_right_number)

        if main_display_width % 2 == 0:
            hori_midpoint = math.floor((colLettersToNumber(bottom_right_letter) - colLettersToNumber(top_left_letter))/2 + colLettersToNumber(top_left_letter))
            north_merge_range_str = colNumberToLetters(hori_midpoint) + str(top_left_number) + ":" + colNumberToLetters(hori_midpoint + 1) + str(top_left_number)
            # print(north_merge_range_str)
            ws.merge_cells(north_merge_range_str)
            north_merge_range = north_merge_range_str.split(':')
            cell_north = ws[north_merge_range[0]]

            south_merge_range_str = relativeToCell(north_merge_range[0], [('down', main_display_height-1)]) + ":" + relativeToCell(north_merge_range[0], [('down', main_display_height-1), ('right', 1)])
            ws.merge_cells(south_merge_range_str)
            south_merge_range = south_merge_range_str.split(":")
            cell_south = ws[south_merge_range[0]]

        else: 
            hori_midpoint = math.floor((colLettersToNumber(bottom_right_letter) - colLettersToNumber(top_left_letter))/2 + colLettersToNumber(top_left_letter))
            cell_north = ws[colNumberToLetters(hori_midpoint) + str(top_left_number)]
            cell_south = ws[relativeToCell(colNumberToLetters(hori_midpoint) + str(top_left_number), [('down', main_display_height-1)])]

        if main_display_height % 2 == 0:
            vert_midpoint = math.floor((bottom_right_number - top_left_number)/2 + top_left_number)
            west_merge_range_str = top_left_letter + str(vert_midpoint) + ":" + top_left_letter + str(vert_midpoint + 1)
            # print(west_merge_range_str)
            ws.merge_cells(west_merge_range_str)
            west_merge_range = west_merge_range_str.split(':')
            cell_west = ws[west_merge_range[0]]

            east_merge_range_str = relativeToCell(west_merge_range[0], [('right', main_display_width-1)]) + ":" + relativeToCell(west_merge_range[0], [('right', main_display_width-1), ('down', 1)])
            ws.merge_cells(east_merge_range_str)
            east_merge_range = east_merge_range_str.split(":")
            cell_east = ws[east_merge_range[0]]

        else: 
            vert_midpoint = math.floor((colLettersToNumber(bottom_right_letter) - colLettersToNumber(top_left_letter))/2 + colLettersToNumber(top_left_letter))
            cell_west = ws[top_left_letter + str(vert_midpoint)]
            cell_east = ws[relativeToCell(top_left_letter + str(vert_midpoint), [('down', main_display_height-1)])]


        cell_north.value = 'N'
        cell_north.alignment = Alignment(horizontal='center', vertical='center')
        cell_north.font = Font(bold=True)
        
        cell_south.value = 'S'
        cell_south.alignment = Alignment(horizontal='center', vertical='center')
        cell_south.font = Font(bold=True)

        cell_west.value = 'W'
        cell_west.alignment = Alignment(horizontal='center', vertical='center')
        cell_west.font = Font(bold=True)
        
        cell_east.value = 'E'
        cell_east.alignment = Alignment(horizontal='center', vertical='center')
        cell_east.font = Font(bold=True)

    createThickOutsideBorders(ws, main_coords[0], main_coords[1])
    createIntersectionBorders(ws, main_coords)

    return None

def importVolumes(ws, df, df_row, origin, travel_dir):
    height = main_display_height
    width = main_display_width

    lanes = []
    for i in range(1,7):
        lanes.append(travel_dir + str(i))

    if travel_dir == 'EB':
        translation_list = [('down', int(height/2 + 3)), ('right', int(width/2 - 4))]
    elif travel_dir == 'NB':
        translation_list = [('down', int(height/2 + 3)), ('right', int(width/2 + 4))]
    elif travel_dir == 'WB':
        translation_list = [('down', int(height/2 - 3)), ('right', int(width/2 + 3))]
    elif travel_dir == 'SB':
        translation_list = [('down', int(height/2 - 4)), ('right', int(width/2 - 4))]

    if header:
        translation_list.append(('down', header_height))

    # print(df.loc[df_row])
    # print(lanes)

    lane_coord = relativeToCell(origin, translation_list)

    for lane_num in lanes: 

        print(df.loc[df_row, 'Scenario'] + ' ' + lane_num)
        print(df.loc[df_row, lane_num])

        print('Lane Coordinate: ' + lane_coord)

        ws[lane_coord] = df.loc[df_row, lane_num]
        
        if travel_dir == 'EB':
            ws[lane_coord].alignment = Alignment(horizontal='right')
            lane_coord = relativeToCell(lane_coord, [('down', 1)])
        elif travel_dir == 'NB':
            cells_to_merge_nb = lane_coord + ":" + relativeToCell(lane_coord, [('down', int(height/2)-5)])
            ws.merge_cells(cells_to_merge_nb)
            ws[lane_coord].alignment = Alignment(textRotation=90, vertical='top')
            lane_coord = relativeToCell(lane_coord, [('right', 1)])
        elif travel_dir == 'WB':
            ws[lane_coord].alignment = Alignment(horizontal='left')
            lane_coord = relativeToCell(lane_coord, [('up', 1)])
        elif travel_dir == 'SB':
            cells_to_merge_sb = relativeToCell(lane_coord, [('up', int(height/2)-5)]) + ":" + lane_coord
            ws.merge_cells(cells_to_merge_sb)
            lane_coord_sb_merged = relativeToCell(lane_coord, [('up', int(height/2)-5)])
            ws[lane_coord_sb_merged].alignment = Alignment(textRotation=90, vertical='bottom')
            # Special case for SB volumes because of the cell merging coordinate change
            ws[lane_coord_sb_merged] = df.loc[df_row, lane_num]

            lane_coord = relativeToCell(lane_coord, [('left', 1)])

            

    

    # print(lanes)
    return

def populateFigure(ws, df, df_row, origin):
    ws[origin] = origin
    height = main_display_height
    width = main_display_width

    # Scenario
    scenario = df.loc[df_row, 'Scenario']
    scenario_cell = relativeToCell(origin, [('right', width + 1)])
    ws[scenario_cell] = scenario
    
    # Road Names [STILL NEEDS TO ACCOMMODATE OVERFLOW TEXT, ALSO DOUBLE CHECK HOW THESE KEYS COME (WHETHER THERE IS A SPACE OR NOT)]
    eb_roadname = df.loc[df_row, 'EB Road Name ']
    eb_translate_list = [('down', int(height/2)), ('right', 1)]
    if header:
        eb_translate_list.append(('down', header_height))
    eb_roadname_cell = relativeToCell(origin, eb_translate_list)
    ws[eb_roadname_cell] = eb_roadname

    wb_roadname = df.loc[df_row, 'WB Road Name']
    wb_translate_list = [('down', int(height/2 - 1)), ('right', width - 2)]
    if header:
        wb_translate_list.append(('down', header_height))
    wb_roadname_cell = relativeToCell(origin, wb_translate_list)
    ws[wb_roadname_cell] = wb_roadname
    ws[wb_roadname_cell].alignment = Alignment(horizontal='right')

    nb_roadname = df.loc[df_row, 'NB Road Name ']
    nb_translate_list_topleft = [('right', int(width/2)-1), ('down', 1)]
    nb_translate_list_bottomright = [('right', int(width/2)-1), ('down', int(height/2-1))]
    if header:
        nb_translate_list_topleft.append(('down', header_height))
        nb_translate_list_bottomright.append(('down', header_height))
    nb_roadname_topleft_cell = relativeToCell(origin, nb_translate_list_topleft)
    nb_roadname_bottomright_cell = relativeToCell(origin, nb_translate_list_bottomright)
    nb_roadname_cellstomerge = nb_roadname_topleft_cell + ":" + nb_roadname_bottomright_cell
    ws.merge_cells(nb_roadname_cellstomerge)
    ws[nb_roadname_topleft_cell] = nb_roadname
    ws[nb_roadname_topleft_cell].alignment = Alignment(textRotation=90, vertical='top')

    nb_roadname = df.loc[df_row, 'NB Road Name ']
    nb_translate_list_topleft = [('right', int(width/2)), ('down', int(height/2))]
    nb_translate_list_bottomright = [('right', int(width/2)), ('down', int(height)-2)]
    if header:
        nb_translate_list_topleft.append(('down', header_height))
        nb_translate_list_bottomright.append(('down', header_height))
    nb_roadname_topleft_cell = relativeToCell(origin, nb_translate_list_topleft)
    nb_roadname_bottomright_cell = relativeToCell(origin, nb_translate_list_bottomright)
    nb_roadname_cellstomerge = nb_roadname_topleft_cell + ":" + nb_roadname_bottomright_cell
    ws.merge_cells(nb_roadname_cellstomerge)
    ws[nb_roadname_topleft_cell] = nb_roadname
    ws[nb_roadname_topleft_cell].alignment = Alignment(textRotation=90, vertical='bottom')

    # Volumes 
    importVolumes(ws, df, df_row, origin, 'EB')
    importVolumes(ws, df, df_row, origin, 'NB')
    importVolumes(ws, df, df_row, origin, 'WB')
    importVolumes(ws, df, df_row, origin, 'SB')
    return

# GET THE CURRENT SCRIPT DIRECTORY
script_dir = os.path.dirname(os.path.abspath(__file__))

# REFERENCE DATA MERGE.CSV INTO DATAFRAME
csv_path = script_dir + "\_data merge.csv"
df = pd.read_csv(csv_path)

df['Scenario'] = df['Scenario'].fillna(method='ffill')

print(df.head(25))
print(df.loc[2])

# GLOBAL VARIABLES FOR OPTIONS
origin = 'B2'
main_display_height = 26    # would recommend keeping this even (code may not work for odd) and at least 26
main_display_width = 24     # would recommend keeping this even (code may not work for odd) and at least 24

header = True            # Boolean for header
if header:
    header_height = 2
main_border = True       # Boolean for border around main area
cardinal_dirs = True     # Boolean for cardinal directions

main_background_color = 'A6C9EC'
main_border_color = 'DAE9F8'      
header_color = '83CCEB'

gap = 3
jump = main_display_height + gap 
if header:
    jump = jump + 2

#  CREATE EXCEL SPREADSHEET
wb = Workbook()

unique_scenarios = df['Scenario'].unique()
for scenario in unique_scenarios: 
    ws = wb.create_sheet(scenario)
    ws.sheet_view.zoomScale = 115

    setColumnWidths(ws, 3)

    origin_col, origin_row = splitCellCoord(origin)
    curr_row = int(origin_row)
    for i in range(len(df.index)):

        if df.loc[i, 'Scenario'] == scenario:
            local_fig_origin = origin_col + str(curr_row)
            generateFigure(ws, local_fig_origin)
            populateFigure(ws, df, i, local_fig_origin)
            curr_row = curr_row + jump


# SAVE
wb.remove(wb['Sheet'])
wb.save(fr'{script_dir}\Figures.xlsx')

