import xlsxwriter
import datetime
import math


def create_excel():
    # Create a workbook and add a worksheet
    global workbook
    global worksheet
    workbook = xlsxwriter.Workbook('FLL Schedule.xlsx')
    worksheet = workbook.add_worksheet()

    # Creates a list containing 25 lists, each of 40 items (representing each col), all set to 0
    global used_cells
    cols, rows = 25, 40
    used_cells = [[0 for x in range(cols)] for y in range(rows)]

    # print(used_cells)
    # used_cells[7][1] = 4
    # print(used_cells)
    # print(used_cells[7][1])


def create_formatting():
    # Create formatting
    global merge_format
    merge_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_name': 'Calibri',
        'font_size': '10',
        'text_wrap': 1
    })

    global merge_format_yellow
    merge_format_yellow = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': '10',
        'fg_color': 'yellow'
    })

    global hours_format
    hours_format = workbook.add_format({
        'num_format': 'hh:mm',
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': '10'
    })

    global hours_format_colored
    hours_format_colored = workbook.add_format({
        'num_format': 'hh:mm',
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': '10',
        'bg_color': '#d6e0f1'
    })

    global row_format
    row_format = workbook.add_format({
        'bg_color': 'white',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': '10'
    })

    global row_format_colored
    row_format_colored = workbook.add_format({
        'bg_color': '#d6e0f1',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': '10'
    })


def format_sheet():
    # Adjust column size
    worksheet.set_column('A:R', 6.57)
    worksheet.set_row(1, 25.5)

    # Creating headings
    worksheet.merge_range('A1:R1', 'FLL ROBOTICS COMPETITION',  merge_format)
    worksheet.merge_range('A2:A3', 'Location',                  merge_format)

    worksheet.merge_range('B2:G2', 'Performance Tables (Gym)',  merge_format)
    worksheet.merge_range('H2:I2', 'Practice Tables (Gym)',     merge_format)

    worksheet.merge_range('J2:L2', 'Robot Design',              merge_format)
    worksheet.merge_range('M2:O2', 'Project Presentation',      merge_format)
    worksheet.merge_range('P2:R2', 'Core Values',               merge_format)

    row = 2
    for col in range(7):
        worksheet.write_number(row, col, col, merge_format)

    worksheet.write_string(2, 7, 'A', merge_format)
    worksheet.write_string(2, 8, 'B', merge_format)

    for col in range(9, 18):
        worksheet.write_number(row, col, col + 206, merge_format)

    # Create times in left column and color rows
    time = datetime.timedelta(hours=8, minutes=30)
    col = 0
    for row in range(3, 38):
        if row % 2 != 0:
            worksheet.write_datetime(row, col, time, hours_format_colored)
        else:
            worksheet.write_datetime(row, col, time, hours_format)
        time += datetime.timedelta(minutes=10)

    # Fill formatting for colored cells throughout schedule
    for row in range(3, 38):
        for col in range(1, 18):
            if row % 2 != 0:
                worksheet.write_blank(row, col, '', row_format_colored)
            else:
                worksheet.write_blank(row, col, '', row_format)

    # Create Breaks in Schedule
    worksheet.merge_range('B4:R6', 'Opening Ceremonies', merge_format_yellow)
    worksheet.merge_range('B7:I7', 'CROSSED CELLS ARE PRACTICE ONLY', merge_format)
    worksheet.merge_range('B16:R16', 'Break', merge_format_yellow)
    worksheet.merge_range('B25:R27', 'Lunch', merge_format_yellow)


def fill_rooms(num_teams):
    # Fill Robot Design rooms for all teams through day
    row = 6
    col = 9
    current_team = 1
    for x in range(0, num_teams):
        if row % 2 == 0:
            worksheet.write_number(row,     col, current_team, row_format)
            worksheet.write_number(row + 1, col, current_team, row_format_colored)
        else:
            worksheet.write_number(row,     col, current_team, row_format_colored)
            worksheet.write_number(row + 1, col, current_team, row_format)

        used_cells[col][row] = current_team
        used_cells[col][row + 1] = current_team

        if row == 6:
            row += 4
        else:
            if col == 11:
                row += 3
                col = 9
            else:
                col += 1
        current_team += 1

    # Fill Project Presentation rooms for all teams through day
    row = 6
    col = 12
    current_team = 4
    for x in range(0, num_teams):
        if row % 2 == 0:
            worksheet.write_number(row,     col, current_team, row_format)
            worksheet.write_number(row + 1, col, current_team, row_format_colored)
        else:
            worksheet.write_number(row,     col, current_team, row_format_colored)
            worksheet.write_number(row + 1, col, current_team, row_format)

        used_cells[col][row] = current_team
        used_cells[col][row + 1] = current_team

        if row == 6:
            row += 4
            col = 12
        else:
            if col == 14:
                row += 3
                col = 12
            else:
                col += 1
        if current_team == num_teams:
            current_team = 0
        current_team += 1

    # Fill Core Values rooms for all teams through day
    row = 6
    col = 15
    current_team = 7
    for x in range(0, num_teams):
        if row % 2 == 0:
            worksheet.write_number(row,     col, current_team, row_format)
            worksheet.write_number(row + 1, col, current_team, row_format_colored)
        else:
            worksheet.write_number(row,     col, current_team, row_format_colored)
            worksheet.write_number(row + 1, col, current_team, row_format)

        used_cells[col][row] = current_team
        used_cells[col][row + 1] = current_team

        if row == 6:
            row += 4
            col = 15
        else:
            if col == 17:
                row += 3
                col = 15
            else:
                col += 1
        if current_team == num_teams:
            current_team = 0
        current_team += 1


def check_range_for_value(xmin, xmax, ymin, ymax, value):
    if xmin == xmax and ymin == ymax:
        if used_cells[xmin][ymin] == value:
            return 0
    if xmin == xmax:
        for y in range(ymin, ymax):
            if used_cells[xmin][y] == value:
                return 0
    for x in range(xmin, xmax):
        if ymin == ymax:
            if used_cells[x][ymin] == value:
                return 0
        else:
            for y in range(ymin, ymax):
                if used_cells[x][y] == value:
                    return 0
    return 1


def check_if_busy_for_practice(num_prac_mats, num_rooms, row, col, team_number):
    # Check if team is in core room 10 min before or after a potential practice to give travel time
    if check_range_for_value(num_prac_mats + 1, num_rooms + num_prac_mats + 1, row - 1, row + 2, team_number):
        # Check if team is already booked for this 10 min
        if check_range_for_value(1, num_prac_mats + 1, row, row, team_number):
            # Check if cell is empty
            if used_cells[col][row] == 0:
                print('they are not busy')
                # print('team #', team_number, 'is not busy for cell', col, row, 'bc the current value is', used_cells[row][col])
                return 1
    return 0


def check_if_possible_for_practice(num_prac_mats, num_rooms, row, col, team_number):
    # check if its possible to place a team in the remaining cells
    for x in range(7, 15):
        for y in range(1, num_prac_mats + 1):
            if used_cells[y][x] == 0:
                if check_if_busy_for_practice(num_prac_mats, num_rooms, x, y, team_number):
                    print('its possible to practice at', y, x)
                    return 1
    return 0


# def swap_practice(num_prac_mats, num_rooms, row, col, team_number, teams):
#     # check if possible to swap the current team which there is no room left for with an earlier team to fit them
#     for x in range(7, 15):
#         for y in range(1, num_prac_mats + 1):
#             if check_if_possible_for_practice(num_prac_mats, num_rooms, y, x, used_cells[y][x]):
#                 print('it is possible to swap', team_number, 'in', y, x, 'with', used_cells[y][x])
#                 teams.append(used_cells[y][x])
#                 teams.remove(team_number)
#                 used_cells[y][x] = team_number
#                 if row % 2 == 0:
#                     worksheet.write_number(row, col, team_number, row_format)
#                 else:
#                     worksheet.write_number(row, col, team_number, row_format_colored)
#                 return 0

def check_if_busy_for_practice_after_lunch(num_prac_mats, num_robot_rooms, row, col, team_number):
    # Check if team is already booked for this 10 min
    if check_range_for_value(1, num_prac_mats + num_robot_rooms + 1, row, row, team_number):
        # Check if cell is empty
        if used_cells[col][row] == 0:
            # print('they are not busy')
            return 1
    return 0

def dump_team(num_prac_mats, num_rooms, teams, num_performance_tables, num_robot_rooms):
    # Practice between break and lunch
    for row in range(16, 24):
        for col in range(num_performance_tables + 1, num_prac_mats + 1):
            for team_number in teams:
                if check_if_busy_for_practice(num_prac_mats, num_rooms, row, col, team_number):
                    teams.remove(team_number)
                    used_cells[col][row] = team_number
                    if row % 2 == 0:
                        worksheet.write_number(row, col, team_number, row_format)
                    else:
                        worksheet.write_number(row, col, team_number, row_format_colored)
                    return
    # Practice after Lunch
    for row in range(27, 31):
        for col in range(num_performance_tables + 1, num_robot_rooms + num_prac_mats + 1):
            for team_number in teams:
                if check_if_busy_for_practice_after_lunch(num_prac_mats, num_robot_rooms, row, col, team_number):
                    teams.remove(team_number)
                    used_cells[col][row] = team_number
                    if row % 2 == 0:
                        worksheet.write_number(row, col, team_number, row_format)
                    else:
                        worksheet.write_number(row, col, team_number, row_format_colored)
                    return


def fill_practice(num_teams, num_performance_tables, num_practice_tables, num_rooms):
    # Fill practice on performance tables and practice tables before the break
    num_prac_mats = num_performance_tables + num_practice_tables

    teams = []
    for j in range(1, num_teams + 1):
        teams.append(j)

    num_robot_rooms = int(num_rooms / 3)
    max_practices = num_prac_mats * 8 / num_teams
    max_practices = math.floor(max_practices)
    num_practices_created = 1

    # while num_practices_created < max_practices:
    for i in range(0, 10):
        # practice before break
        past_teams = teams
        for row in range(7, 15):
            for col in range(1, num_prac_mats + 1):
                for team_number in teams:
                    print(num_prac_mats, num_rooms, row, col, team_number)
                    if not check_if_possible_for_practice(num_prac_mats, num_rooms, row, col, team_number):
                        dump_team(num_prac_mats, num_rooms, teams, num_performance_tables, num_robot_rooms)
                    else:
                        print(i)
                        if check_if_busy_for_practice(num_prac_mats, num_rooms, row, col, team_number):
                            print(i)
                            teams.remove(team_number)
                            print('writing', team_number, 'to', col, row)
                            used_cells[col][row] = team_number
                            print('after right the new value in matrix is', used_cells[col][row])
                            if row % 2 == 0:
                                worksheet.write_number(row, col, team_number, row_format)
                            else:
                                worksheet.write_number(row, col, team_number, row_format_colored)
                    break
        if not teams:
            print('making a new teams')
            num_practices_created += 1
            for j in range(1, num_teams + 1):
                teams.append(j)
        # if past_teams == teams:
        #     # This means we must be stuck and can't fit a team into the current spots, so we dump them later in the day
        #     # Find the first white space and dump them there
        #     first_col = 0
        #     first_row = 0
        #     for row in range(16, 24):
        #         for col in range(num_performance_tables + 1, num_prac_mats + 1):
        #             if used_cells[col][row] == 0:
        #                 if first_col == 0:
        #                     first_col = col
        #                     first_row = row
        #     used_cells[first_col][first_row] = teams[0]
        #     worksheet.write_number(first_row, first_col, teams[0], row_format)


    print('created', num_practices_created, 'practices')


def main():
    create_excel()
    create_formatting()
    format_sheet()
    fill_rooms(16)
    fill_practice(16, 6, 2, 9)
    workbook.close()
    print(used_cells)

if __name__ == '__main__':
    main()
