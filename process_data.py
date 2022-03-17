import csv
import os
import xlsxwriter
from dataclasses import dataclass, field


####################
# GLOBAL VARIABLES #
####################


input_file_name = "input.csv"
output_file_name = "output_data.xlsx"

all_team_match_entries = []
all_team_data = []
team_num_list = []
output_worksheets = []
all_team_avg_match_contribution = []
all_team_avg_auto = []
all_team_avg_tele = []
all_team_avg_climb = []
all_team_defense_percent = []
top_teams_across_categories = []

MAX_NUMBER_OF_QUAL_MATCHES = 15
DATA_START_ROW = 37
AVERAGES_ROW = 33
STATISTICS_START_ROW = 16
STATISTICS_START_COL = 0
CHART_START_ROW = 1
CHART_ROW_SPACING = 16
FIRST_CHART_COL = "A"
SECOND_CHART_COL = "E"
THIRD_CHART_COL = "I"
NUM_OF_TOP_TEAMS_TO_COLOR_PER_CATEGORY = 5
current_match_count = 0

# For a single robot
MAX_POSSIBLE_AUTO_POINTS = 10
MAX_POSSIBLE_TELE_POINTS = 30


#####################
# CLASSES AND DICTS #
#####################

chart_colors = {
    "RED": "#EA5545",
    "BLUE": "#27AEEF",
    "YELLOW": "#FFCA3A",
    "GREEN": "#4DA167",
    "PURPLE": "#805D93",
    "BLACK": "#252323",
}

team_labeling_colors = {
    1: "#fe3b1e",
    2: "#f4ea5c",
    3: "#2d69cb",
    4: "#11963b",
    5: "#e61cf7",
    6: "#b7c0ff",
    7: "#ae6507",
    8: "#4f02ec",
    9: "#fa2f7a",
    10: "#51e113",
    11: "#6febff",
    12: "#f7aa30",
    13: "#992f7c",
    14: "#00a6ee",
    15: "#acbe9c",
    16: "#a12c32",
    17: "#9b9500",
    18: "#08fdcc",
    19: "#827c70",
    20: "#8e7ba4",
    21: "#566204",
    22: "#fb9fda",
    23: "#08a29a",
    24: "#2a666a",
    25: "#805D93",
}

team_labeling_formats = {}


taxi_completed_dict = {
    "Yes": 1,
    "No": 0,
    "": -1,
}

hangar_level_points_dict = {
    "No Hang": 0,
    "Low Rung (1)": 4,
    "Mid Rung (2)": 6,
    "High Rung (3)": 10,
    "Traversal Rung (4)": 15,
    "": -1
}

hangar_level_empty_count_dict = {
    "No Hang": 0,
    "Low Rung (1)": 0,
    "Mid Rung (2)": 0,
    "High Rung (3)": 0,
    "Traversal Rung (4)": 0,
}

defense_level_as_percentage_dict = {
    "No": 0,
    "Unsure": 0.5,
    "Yes": 1,
    "": -1,
}

defense_level_empty_count_dict = {
    "No": 0,
    "Unsure": 0,
    "Yes": 0,
}


@dataclass
class SingleTeamSingleMatchEntry:
    """Class to keep track of one team's performance during one match"""
    team_num: int
    qual_match_num: int
    successfully_completed_taxi: bool
    auto_cargo_scored_upper: int
    auto_cargo_scored_lower: int
    tele_cargo_scored_upper: int
    tele_cargo_scored_lower: int
    hangar_level: int
    defense_level: int
    other_info: str


@dataclass
class TeamData:
    team_num: int = 0
    match_data: list = field(default_factory=list)
    taxi_percent: float = 0
    avg_auto_points: float = 0
    avg_tele_points: float = 0
    avg_defense_equivalent: float = 0
    avg_climb_points: float = 0
    hangar_level_count_dict: dict = field(default_factory=dict)
    defense_level_count_dict: dict = field(default_factory=dict)


#############
# FUNCTIONS #
#############


def get_max_value_from_comma_separated_numbers(numbers):
    # The argument "numbers" comes in as a string of numbers
    # e.x. "0, 1, 2, 3" or "0, 1" or just a single number as a string like "3"
    number_str_list = numbers.split(",")
    number_int_list = []

    for num in number_str_list:
        # Sometimes the split() function above leaves blank entries, this handles that issue
        if num != "":
            number_int_list.append(int(num))
        else:
            number_int_list.append(0)

    # Just in case number_int_list is empty, add -1 so that it will be the max number and
    # indicate that there was an error
    number_int_list.append(-1)

    # Get the max value from the list of numbers
    max_number = max(number_int_list)

    return max_number


def parse_team_number(num):
    # The team number could come in as a poorly formatted string or a number, this function
    # helps standardize the input
    if type(num) == int:
        parsed_num = num
    elif type(num) == str:
        if num != "":
            parsed_num = int(float(num))
        else:
            parsed_num = -1

    return parsed_num


def parse_match_number(num):
    # The match number could come in as a poorly formatted string or a number, this function
    # helps standardize the input
    if type(num) == int:
        parsed_num = num
    elif type(num) == str:
        if num != "":
            parsed_num = int(float(num))
        else:
            parsed_num = -1

    return parsed_num


######################
# PROCESS INPUT DATA #
######################


# Open the input data spreadsheet and call it input_worksheet while inside of the "with"
# statement
with open(input_file_name, "r", newline="") as input_csv_file:

    # Get rid of any existing output file
    if os.path.exists(output_file_name):
        os.remove(output_file_name)

    # Create the csv handling object out of the open csv file
    input_handling_object = csv.reader(input_csv_file)

    for row_num, row_data in enumerate(input_handling_object):
        # Skip the 0th row (column titles). Python starts counting at 0 instead of 1
        if row_num > 0:
            # Put all of the information from a single row in excel into a python object
            # called "team_match_entry" to make it easier to deal with later on
            print(f"Processing Row Number {row_num + 1}")
            team_match_entry = SingleTeamSingleMatchEntry(
                team_num=parse_team_number(row_data[3]),
                qual_match_num=parse_match_number(row_data[4]),
                successfully_completed_taxi=taxi_completed_dict[row_data[5]],
                auto_cargo_scored_upper=get_max_value_from_comma_separated_numbers(
                    row_data[6]),
                auto_cargo_scored_lower=get_max_value_from_comma_separated_numbers(
                    row_data[7]),
                tele_cargo_scored_upper=get_max_value_from_comma_separated_numbers(
                    row_data[8]),
                tele_cargo_scored_lower=get_max_value_from_comma_separated_numbers(
                    row_data[9]),
                hangar_level=hangar_level_points_dict[row_data[10]],
                defense_level=defense_level_as_percentage_dict[row_data[11]],
                other_info=row_data[12]
            )

            # Add the single-team single-match entry (i.e. data from one row) to a list
            # containing all of these entries
            all_team_match_entries.append(team_match_entry)
        else:
            print("Skipping Row Number 1 (Column Titles)")


# Go through every match entry one-by-one and check if a class for all of a team's matches
# has been created yet (the TeamData class). If not, create it and add it to a list of
# these classes. Then, add the current match entry to its corresponding class containing all
# of that team's match entries.
for match_entry in all_team_match_entries:
    # If the team num is -1 (due to an input error), skip this iteration of the for loop
    if match_entry.team_num == -1:
        continue

    if match_entry.team_num not in team_num_list:
        team_num_list.append(match_entry.team_num)

        new_single_team_data = TeamData()
        new_single_team_data.team_num = match_entry.team_num
        new_single_team_data.hangar_level_count_dict = {
            key: value for key, value in hangar_level_empty_count_dict.items()}
        new_single_team_data.defense_level_count_dict = {
            key: value for key, value in defense_level_empty_count_dict.items()}

        all_team_data.append(new_single_team_data)

    for single_teams_data in all_team_data:
        if single_teams_data.team_num == match_entry.team_num:

            duplicate_entry = False

            for single_team_match in single_teams_data.match_data:
                if match_entry.qual_match_num == single_team_match.qual_match_num:
                    duplicate_entry = True

            if not duplicate_entry:
                single_teams_data.match_data.append(match_entry)


###############################
# GENERATE OUTPUT SPREADSHEET #
###############################


# Create a new workbook for the nicely formatted output workbook with each team as a
# separate tab
with xlsxwriter.Workbook(output_file_name) as output_workbook:

    # Cell formatting objects to format cells as percents, decimals, etc.
    percent_format = output_workbook.add_format({'num_format': '0.0%'})
    one_decimal_format = output_workbook.add_format({'num_format': '0.0'})

    # Create color formats for each team color for use later on in the ranking sheet
    for key, value in team_labeling_colors.items():
        new_format = output_workbook.add_format({'num_format': '0'})
        new_format.set_bg_color(value)
        team_labeling_formats.update({key: new_format})

    # Sort the team number list in ascending order
    team_num_list = sorted(team_num_list)

    # Create the ranking worksheet first
    ranking_worksheet = output_workbook.add_worksheet("Rankings")
    output_worksheets.append(ranking_worksheet)

    # Create a new sheet for every team
    for team_num in team_num_list:
        single_teams_worksheet = output_workbook.add_worksheet(str(team_num))
        output_worksheets.append(single_teams_worksheet)

        # Find the all_team_data entry for the team with the same number as the team_num
        # variable
        for single_teams_data in all_team_data:
            if single_teams_data.team_num == team_num:

                # Sort the match data to be in ascending order of qual match number. I don't
                # fully understand how this works, but I got it from stack overflow and it
                # does the job.
                single_teams_data.match_data = sorted(
                    single_teams_data.match_data,
                    key=lambda x: x.qual_match_num,
                    reverse=False)

                # Populate the worksheet for this team with match data, graphs, etc.

                # Data category titles for averages
                single_teams_worksheet.write(
                    AVERAGES_ROW - 1, 1, "Taxi")
                single_teams_worksheet.write(
                    AVERAGES_ROW - 1, 2, "AUTO - Cargo Scored [Upper Hub]")
                single_teams_worksheet.write(
                    AVERAGES_ROW - 1, 3, "AUTO - Cargo Scored [Lower Hub]")
                single_teams_worksheet.write(
                    AVERAGES_ROW - 1, 4, "TELEOP - Cargo Scored [Upper Hub]")
                single_teams_worksheet.write(
                    AVERAGES_ROW - 1, 5, "TELEOP - Cargo Scored [Lower Hub]")
                single_teams_worksheet.write(
                    AVERAGES_ROW - 1, 6, "Hangar")
                single_teams_worksheet.write(
                    AVERAGES_ROW - 1, 7, "Mostly defense?")

                # Data category titles for match data
                single_teams_worksheet.write(
                    DATA_START_ROW - 1, 0, "MATCH DATA")
                single_teams_worksheet.write(
                    DATA_START_ROW, 0, "Qualification Number")
                single_teams_worksheet.write(
                    DATA_START_ROW, 1, "Taxi")
                single_teams_worksheet.write(
                    DATA_START_ROW, 2, "AUTO - Cargo Scored [Upper Hub]")
                single_teams_worksheet.write(
                    DATA_START_ROW, 3, "AUTO - Cargo Scored [Lower Hub]")
                single_teams_worksheet.write(
                    DATA_START_ROW, 4, "TELEOP - Cargo Scored [Upper Hub]")
                single_teams_worksheet.write(
                    DATA_START_ROW, 5, "TELEOP - Cargo Scored [Lower Hub]")
                single_teams_worksheet.write(
                    DATA_START_ROW, 6, "Hangar")
                single_teams_worksheet.write(
                    DATA_START_ROW, 7, "Mostly defense?")
                single_teams_worksheet.write(
                    DATA_START_ROW, 8, "Other Information")

                # For charts later on
                single_teams_worksheet.write(
                    DATA_START_ROW, 23, "Taxi Num")
                single_teams_worksheet.write(
                    DATA_START_ROW, 24, "Climb Num")
                single_teams_worksheet.write(
                    DATA_START_ROW, 25, "Defense Num")

                # Set the column widths to make the text legible
                single_teams_worksheet.set_column_pixels(0, 0, 120)   # Qual
                single_teams_worksheet.set_column_pixels(1, 1, 50)   # Taxi
                single_teams_worksheet.set_column_pixels(
                    2, 3, 180)                                        # Auto cargo (both)
                single_teams_worksheet.set_column_pixels(
                    4, 5, 190)                                        # Tele cargo (both)
                single_teams_worksheet.set_column_pixels(6, 6, 100)    # Hangar
                single_teams_worksheet.set_column_pixels(7, 7, 90)    # Defense
                single_teams_worksheet.set_column_pixels(8, 8, 1000)  # Other

                # These variables keep track of the total points scored in each category
                # so that averages can be calculated later
                team_total_taxi_equivalent = 0
                team_total_auto_cargo_upper = 0
                team_total_auto_cargo_lower = 0
                team_total_tele_cargo_upper = 0
                team_total_tele_cargo_lower = 0
                team_total_defense_equivalent = 0
                team_total_climb_points = 0

                # Add the team's match data to that team's worksheet
                for i, match in enumerate(single_teams_data.match_data):
                    # Turn the python representation of the match data back into
                    # human-friendly text that can be written to the spreadsheet
                    taxi_string = "Yes" if match.successfully_completed_taxi == 1 else "No"

                    hangar_string = "ERROR"
                    for key, value in hangar_level_points_dict.items():
                        if match.hangar_level == value:
                            hangar_string = key
                            single_teams_data.hangar_level_count_dict[key] += 1

                    defense_string = "ERROR"
                    for key, value in defense_level_as_percentage_dict.items():
                        if match.defense_level == value:
                            defense_string = key
                            single_teams_data.defense_level_count_dict[key] += 1

                    # i + 1 is needed because i starts counting at 0, but I want the first
                    # row of match data to be in the row at index 1, not 0.
                    single_teams_worksheet.write(
                        DATA_START_ROW + i + 1, 0, match.qual_match_num)
                    single_teams_worksheet.write(
                        DATA_START_ROW + i + 1, 1, taxi_string)
                    single_teams_worksheet.write(
                        DATA_START_ROW + i + 1, 2, match.auto_cargo_scored_upper)
                    single_teams_worksheet.write(
                        DATA_START_ROW + i + 1, 3, match.auto_cargo_scored_lower)
                    single_teams_worksheet.write(
                        DATA_START_ROW + i + 1, 4, match.tele_cargo_scored_upper)
                    single_teams_worksheet.write(
                        DATA_START_ROW + i + 1, 5, match.tele_cargo_scored_lower)
                    single_teams_worksheet.write(
                        DATA_START_ROW + i + 1, 6, hangar_string)
                    single_teams_worksheet.write(
                        DATA_START_ROW + i + 1, 7, defense_string)
                    single_teams_worksheet.write(
                        DATA_START_ROW + i + 1, 8, match.other_info)

                    # For charts later on
                    single_teams_worksheet.write(
                        DATA_START_ROW + i + 1, 23, match.successfully_completed_taxi)
                    single_teams_worksheet.write(
                        DATA_START_ROW + i + 1, 24, match.hangar_level)
                    single_teams_worksheet.write(
                        DATA_START_ROW + i + 1, 25, match.defense_level)

                    # Add up the team total numbers across all matches
                    team_total_taxi_equivalent += match.successfully_completed_taxi
                    team_total_auto_cargo_upper += match.auto_cargo_scored_upper
                    team_total_auto_cargo_lower += match.auto_cargo_scored_lower
                    team_total_tele_cargo_upper += match.tele_cargo_scored_upper
                    team_total_tele_cargo_lower += match.tele_cargo_scored_lower
                    team_total_defense_equivalent += match.defense_level
                    team_total_climb_points += match.hangar_level

                # Write climb total counts to sheet for pie charts later
                single_teams_worksheet.write(
                    DATA_START_ROW, 19, "Hangar levels")
                single_teams_worksheet.write(
                    DATA_START_ROW, 20, "Frequency")
                for i, (key, value) in enumerate(single_teams_data.hangar_level_count_dict.items()):
                    single_teams_worksheet.write(
                        DATA_START_ROW + i + 1, 19, key)
                    single_teams_worksheet.write(
                        DATA_START_ROW + i + 1, 20, value)

                # Write defense total counts to sheet for pie charts later
                single_teams_worksheet.write(
                    DATA_START_ROW, 21, "Hangar levels")
                single_teams_worksheet.write(
                    DATA_START_ROW, 22, "Frequency")
                for i, (key, value) in enumerate(single_teams_data.defense_level_count_dict.items()):
                    single_teams_worksheet.write(
                        DATA_START_ROW + i + 1, 21, key)
                    single_teams_worksheet.write(
                        DATA_START_ROW + i + 1, 22, value)

                # Get the current number of matches this team has completed
                current_match_count = len(single_teams_data.match_data)

                # Create averages from totals
                team_avg_taxi_percent = team_total_taxi_equivalent / current_match_count
                team_avg_auto_cargo_upper = team_total_auto_cargo_upper / current_match_count
                team_avg_auto_cargo_lower = team_total_auto_cargo_lower / current_match_count
                team_avg_tele_cargo_upper = team_total_tele_cargo_upper / current_match_count
                team_avg_tele_cargo_lower = team_total_tele_cargo_lower / current_match_count
                team_avg_defense_equivalent = team_total_defense_equivalent / current_match_count
                team_avg_climb_points = team_total_climb_points / current_match_count

                # Update general team statistics based on these averages
                single_teams_data.taxi_percent = team_avg_taxi_percent
                single_teams_data.avg_auto_points = 2 * team_avg_taxi_percent + \
                    2 * team_avg_auto_cargo_lower + 4 * team_avg_auto_cargo_upper
                single_teams_data.avg_tele_points = team_avg_tele_cargo_lower + \
                    2 * team_avg_auto_cargo_upper
                single_teams_data.avg_defense_equivalent = team_avg_defense_equivalent
                single_teams_data.avg_climb_points = team_avg_climb_points

                # Add team statistics to lists for ranking
                all_team_avg_match_contribution.append(
                    (team_num, single_teams_data.avg_auto_points + single_teams_data.avg_tele_points + single_teams_data.avg_tele_points))
                all_team_avg_auto.append(
                    (team_num, single_teams_data.avg_auto_points))
                all_team_avg_tele.append(
                    (team_num, single_teams_data.avg_tele_points))
                all_team_avg_climb.append(
                    (team_num, single_teams_data.avg_climb_points))
                all_team_defense_percent.append(
                    (team_num, single_teams_data.avg_defense_equivalent))

                # Print averages to the spreadsheet
                single_teams_worksheet.write(
                    AVERAGES_ROW, 0, "Averages:")
                single_teams_worksheet.write(
                    AVERAGES_ROW, 1, team_avg_taxi_percent, percent_format)
                single_teams_worksheet.write(
                    AVERAGES_ROW, 2, team_avg_auto_cargo_upper, one_decimal_format)
                single_teams_worksheet.write(
                    AVERAGES_ROW, 3, team_avg_auto_cargo_lower, one_decimal_format)
                single_teams_worksheet.write(
                    AVERAGES_ROW, 4, team_avg_tele_cargo_upper, one_decimal_format)
                single_teams_worksheet.write(
                    AVERAGES_ROW, 5, team_avg_tele_cargo_lower, one_decimal_format)
                single_teams_worksheet.write(
                    AVERAGES_ROW, 6, team_avg_climb_points, one_decimal_format)
                single_teams_worksheet.write(
                    AVERAGES_ROW, 7, team_avg_defense_equivalent, percent_format)

                # Summary statistics
                single_teams_worksheet.write(
                    STATISTICS_START_ROW,
                    STATISTICS_START_COL,
                    "Taxi Percentage: ")
                single_teams_worksheet.write(
                    STATISTICS_START_ROW,
                    STATISTICS_START_COL + 1,
                    single_teams_data.taxi_percent,
                    percent_format)
                single_teams_worksheet.write(
                    STATISTICS_START_ROW + 1,
                    STATISTICS_START_COL,
                    "Avg. Auto Points: ")
                single_teams_worksheet.write(
                    STATISTICS_START_ROW + 1,
                    STATISTICS_START_COL + 1,
                    single_teams_data.avg_auto_points,
                    one_decimal_format)
                single_teams_worksheet.write(
                    STATISTICS_START_ROW + 1,
                    STATISTICS_START_COL + 2,
                    "(Including avg. taxi points)")
                single_teams_worksheet.write(
                    STATISTICS_START_ROW + 2,
                    STATISTICS_START_COL,
                    "Avg. Teleop Points: ")
                single_teams_worksheet.write(
                    STATISTICS_START_ROW + 2,
                    STATISTICS_START_COL + 1,
                    single_teams_data.avg_tele_points,
                    one_decimal_format)
                single_teams_worksheet.write(
                    STATISTICS_START_ROW + 3,
                    STATISTICS_START_COL,
                    "Avg. Climb Points: ")
                single_teams_worksheet.write(
                    STATISTICS_START_ROW + 3,
                    STATISTICS_START_COL + 1,
                    single_teams_data.avg_climb_points,
                    one_decimal_format)
                single_teams_worksheet.write(
                    STATISTICS_START_ROW + 4,
                    STATISTICS_START_COL,
                    "Defense Percentage: ")
                single_teams_worksheet.write(
                    STATISTICS_START_ROW + 4,
                    STATISTICS_START_COL + 1,
                    single_teams_data.avg_defense_equivalent,
                    percent_format)
                single_teams_worksheet.write(
                    STATISTICS_START_ROW + 4,
                    STATISTICS_START_COL + 2,
                    "(100% = Yes/Always, 0% = No/Never)")

                # Create the chart for cargo scored in auto (high vs. low)
                cargo_in_auto_chart = output_workbook.add_chart(
                    {'type': 'column'})
                cargo_in_auto_chart.set_title(
                    {'name': 'AUTO - Upper vs. Lower Hub Cargo'})
                cargo_in_auto_chart.set_x_axis(
                    {'name': 'Qualification Match'})
                cargo_in_auto_chart.set_y_axis(
                    {'name': 'Cargo Scored',
                     'min': 0,
                     'max': MAX_POSSIBLE_AUTO_POINTS})
                cargo_in_auto_chart.add_series({
                    'name': 'Upper Hub',
                    'categories': f'={single_teams_data.team_num}!A{DATA_START_ROW + 2}:A{DATA_START_ROW + current_match_count + 1}',
                    'values': f'={single_teams_data.team_num}!C{DATA_START_ROW + 2}:C{DATA_START_ROW + current_match_count + 1}',
                    'fill': {'color': chart_colors["BLUE"]},
                })
                cargo_in_auto_chart.add_series({
                    'name': 'Lower Hub',
                    'categories': f'={single_teams_data.team_num}!A{DATA_START_ROW + 2}:A{DATA_START_ROW + current_match_count + 1}',
                    'values': f'={single_teams_data.team_num}!D{DATA_START_ROW + 2}:D{DATA_START_ROW + current_match_count + 1}',
                    'fill': {'color': chart_colors["RED"]},
                })

                single_teams_worksheet.insert_chart(
                    f"{FIRST_CHART_COL}{CHART_START_ROW}", cargo_in_auto_chart)

                # Create the chart for cargo scored in teleop (high vs. low)
                cargo_in_teleop_chart = output_workbook.add_chart(
                    {'type': 'column'})
                cargo_in_teleop_chart.set_title(
                    {'name': 'TELEOP - Upper vs. Lower Hub Cargo'})
                cargo_in_teleop_chart.set_x_axis(
                    {'name': 'Qualification Match'})
                cargo_in_teleop_chart.set_y_axis(
                    {'name': 'Cargo Scored',
                     'min': 0,
                     'max': MAX_POSSIBLE_TELE_POINTS})
                cargo_in_teleop_chart.add_series({
                    'name': 'Upper Hub',
                    'categories': f'={single_teams_data.team_num}!A{DATA_START_ROW + 2}:A{DATA_START_ROW + current_match_count + 1}',
                    'values': f'={single_teams_data.team_num}!E{DATA_START_ROW + 2}:E{DATA_START_ROW + current_match_count + 1}',
                    'fill': {'color': chart_colors["BLUE"]},
                })
                cargo_in_teleop_chart.add_series({
                    'name': 'Lower Hub',
                    'categories': f'={single_teams_data.team_num}!A{DATA_START_ROW + 2}:A{DATA_START_ROW + current_match_count + 1}',
                    'values': f'={single_teams_data.team_num}!F{DATA_START_ROW + 2}:F{DATA_START_ROW + current_match_count + 1}',
                    'fill': {'color': chart_colors["RED"]},
                })

                single_teams_worksheet.insert_chart(
                    f"{SECOND_CHART_COL}{CHART_START_ROW}", cargo_in_teleop_chart)

                # Create the chart for hangar level across matches
                hangar_points_over_time_chart = output_workbook.add_chart(
                    {'type': 'column'})
                hangar_points_over_time_chart.set_title(
                    {'name': 'Hangar Points Over Time'})
                hangar_points_over_time_chart.set_x_axis(
                    {'name': 'Qualification Match'})
                hangar_points_over_time_chart.set_y_axis(
                    {'name': 'Hangar Points',
                     'min': 0,
                     'max': 15})
                hangar_points_over_time_chart.add_series({
                    'name': 'Hangar Points',
                    'categories': f'={single_teams_data.team_num}!A{DATA_START_ROW + 2}:A{DATA_START_ROW + current_match_count + 1}',
                    'values': f'={single_teams_data.team_num}!Y{DATA_START_ROW + 2}:Y{DATA_START_ROW + current_match_count + 1}',
                    'fill': {'color': chart_colors["RED"]},
                })

                single_teams_worksheet.insert_chart(
                    f"{SECOND_CHART_COL}{CHART_START_ROW + CHART_ROW_SPACING}", hangar_points_over_time_chart)

                # Hangar Pie Chart
                hangar_pie_chart = output_workbook.add_chart(
                    {'type': 'pie'})
                hangar_pie_chart.set_title(
                    {'name': 'Climb Level Breakdown'})
                hangar_pie_chart.add_series({
                    'name': 'Climb Level',
                    'categories': f'={single_teams_data.team_num}!T{DATA_START_ROW + 2}:T{DATA_START_ROW + 6}',
                    'values': f'={single_teams_data.team_num}!U{DATA_START_ROW + 2}:U{DATA_START_ROW + 6}',
                    'points': [
                        {'fill': {'color': chart_colors["BLACK"]}},
                        {'fill': {'color': chart_colors["RED"]}},
                        {'fill': {'color': chart_colors["YELLOW"]}},
                        {'fill': {'color': chart_colors["BLUE"]}},
                        {'fill': {'color': chart_colors["GREEN"]}},
                    ]
                })

                single_teams_worksheet.insert_chart(
                    f"{THIRD_CHART_COL}{CHART_START_ROW + CHART_ROW_SPACING}", hangar_pie_chart)

                # Defense Pie Chart
                defense_pie_chart = output_workbook.add_chart(
                    {'type': 'pie'})
                defense_pie_chart.set_title(
                    {'name': 'Does This Team Play Defense?'})
                defense_pie_chart.add_series({
                    'name': 'Defense',
                    'categories': f'={single_teams_data.team_num}!V{DATA_START_ROW + 2}:V{DATA_START_ROW + 4}',
                    'values': f'={single_teams_data.team_num}!W{DATA_START_ROW + 2}:W{DATA_START_ROW + 4}',
                    'points': [
                        {'fill': {'color': chart_colors["RED"]}},
                        {'fill': {'color': chart_colors["YELLOW"]}},
                        {'fill': {'color': chart_colors["GREEN"]}},
                    ]
                })

                single_teams_worksheet.insert_chart(
                    f"{THIRD_CHART_COL}{CHART_START_ROW}", defense_pie_chart)

                # TODO: Extremely fancy graphs that look absurd

    # Sort/rank all of the team's individual statistics to find leaders in each category
    all_team_avg_match_contribution = sorted(
        all_team_avg_match_contribution,
        key=lambda x: x[1],
        reverse=True)
    all_team_avg_auto = sorted(
        all_team_avg_auto,
        key=lambda x: x[1],
        reverse=True)
    all_team_avg_tele = sorted(
        all_team_avg_tele,
        key=lambda x: x[1],
        reverse=True)
    all_team_avg_climb = sorted(
        all_team_avg_climb,
        key=lambda x: x[1],
        reverse=True)
    all_team_defense_percent = sorted(
        all_team_defense_percent,
        key=lambda x: x[1],
        reverse=True)

    # Add the top 5 from each category to a set of all of these teams, where duplicates will
    # be avoided. This list will determine which teams get colored
    top_teams_across_categories += all_team_avg_match_contribution[0:
                                                                   NUM_OF_TOP_TEAMS_TO_COLOR_PER_CATEGORY]
    top_teams_across_categories += all_team_avg_auto[0:
                                                     NUM_OF_TOP_TEAMS_TO_COLOR_PER_CATEGORY]
    top_teams_across_categories += all_team_avg_tele[0:
                                                     NUM_OF_TOP_TEAMS_TO_COLOR_PER_CATEGORY]
    top_teams_across_categories += all_team_avg_climb[0:
                                                      NUM_OF_TOP_TEAMS_TO_COLOR_PER_CATEGORY]
    top_teams_across_categories += all_team_defense_percent[0:
                                                            NUM_OF_TOP_TEAMS_TO_COLOR_PER_CATEGORY]

    # Get a list of unique teams that are at the top of at least one category
    top_teams_across_categories = list(set([
        team for (team, junk) in top_teams_across_categories]))

    # Add a number to each team that corresponds to its color on the ranking sheet
    top_teams_across_categories = {
        team: i + 1 for i, team in enumerate(top_teams_across_categories)}

    # Write the column headers
    ranking_worksheet.write(0, 0, "#")
    ranking_worksheet.write(0, 2, "Team")
    ranking_worksheet.write(0, 3, "Avg. Match Contribution (Pts.)")
    ranking_worksheet.write(0, 5, "Team")
    ranking_worksheet.write(0, 6, "Avg. Auto Pts.")
    ranking_worksheet.write(0, 8, "Team")
    ranking_worksheet.write(0, 9, "Avg. Teleop Pts.")
    ranking_worksheet.write(0, 11, "Team")
    ranking_worksheet.write(0, 12, "Avg. Climb Pts.")
    ranking_worksheet.write(0, 14, "Team")
    ranking_worksheet.write(0, 15, "Defense %")

    # Set the column widths so the text is legible
    ranking_worksheet.set_column_pixels(0, 0, 20)  # Num
    ranking_worksheet.set_column_pixels(1, 1, 10)    # Gap
    ranking_worksheet.set_column_pixels(2, 2, 40)   # Team
    ranking_worksheet.set_column_pixels(3, 3, 170)  # Match pts
    ranking_worksheet.set_column_pixels(4, 4, 10)   # Gap
    ranking_worksheet.set_column_pixels(5, 5, 40)   # Team
    ranking_worksheet.set_column_pixels(6, 6, 80)  # Auto pts
    ranking_worksheet.set_column_pixels(7, 7, 10)   # Gap
    ranking_worksheet.set_column_pixels(8, 8, 40)   # Team
    ranking_worksheet.set_column_pixels(9, 9, 90)  # Teleop pts
    ranking_worksheet.set_column_pixels(10, 10, 10)  # Gap
    ranking_worksheet.set_column_pixels(11, 11, 40)  # Team
    ranking_worksheet.set_column_pixels(12, 12, 85)    # Climb
    ranking_worksheet.set_column_pixels(13, 13, 10)  # Gap
    ranking_worksheet.set_column_pixels(14, 14, 40)  # Team
    ranking_worksheet.set_column_pixels(15, 15, 60)  # Defense

    # Add ranking numbers to the rank sheet
    for i, team_num in enumerate(team_num_list):
        ranking_worksheet.write(i + 1, 0, i + 1)

    # List all teams by average match contribution
    for i, (team_num, pts) in enumerate(all_team_avg_match_contribution):
        if team_num in top_teams_across_categories:
            for top_team_num, j in top_teams_across_categories.items():
                if team_num == top_team_num:
                    ranking_worksheet.write(
                        i + 1, 2, team_num, team_labeling_formats[j])
        else:
            ranking_worksheet.write(i + 1, 2, team_num)
        ranking_worksheet.write(i + 1, 3, pts, one_decimal_format)

    # List all teams by average auto points
    for i, (team_num, pts) in enumerate(all_team_avg_auto):
        if team_num in top_teams_across_categories:
            for top_team_num, j in top_teams_across_categories.items():
                if team_num == top_team_num:
                    ranking_worksheet.write(
                        i + 1, 5, team_num, team_labeling_formats[j])
        else:
            ranking_worksheet.write(i + 1, 5, team_num)
        ranking_worksheet.write(i + 1, 6, pts, one_decimal_format)

    # List all teams by average teleop points
    for i, (team_num, pts) in enumerate(all_team_avg_tele):
        if team_num in top_teams_across_categories:
            for top_team_num, j in top_teams_across_categories.items():
                if team_num == top_team_num:
                    ranking_worksheet.write(
                        i + 1, 8, team_num, team_labeling_formats[j])
        else:
            ranking_worksheet.write(i + 1, 8, team_num)
        ranking_worksheet.write(i + 1, 9, pts, one_decimal_format)

    # List all teams by average climb points
    for i, (team_num, pts) in enumerate(all_team_avg_climb):
        if team_num in top_teams_across_categories:
            for top_team_num, j in top_teams_across_categories.items():
                if team_num == top_team_num:
                    ranking_worksheet.write(
                        i + 1, 11, team_num, team_labeling_formats[j])
        else:
            ranking_worksheet.write(i + 1, 11, team_num)
        ranking_worksheet.write(i + 1, 12, pts, one_decimal_format)

    # List all teams by defense percentage
    for i, (team_num, pts) in enumerate(all_team_defense_percent):
        if team_num in top_teams_across_categories:
            for top_team_num, j in top_teams_across_categories.items():
                if team_num == top_team_num:
                    ranking_worksheet.write(
                        i + 1, 14, team_num, team_labeling_formats[j])
        else:
            ranking_worksheet.write(i + 1, 14, team_num)
        ranking_worksheet.write(i + 1, 15, pts, one_decimal_format)

    # Add rank numbers back into each team's worksheet
    for single_team_worksheet in output_worksheets:
        for team_num in team_num_list:
            if str(team_num) == single_team_worksheet.name:
                # Write all of the ranked category titles to the sheet
                single_team_worksheet.write(22, 0, "Ranked Category")
                single_team_worksheet.write(22, 1, "Rank")
                single_team_worksheet.write(23, 0, "Avg. Match Points")
                single_team_worksheet.write(24, 0, "Avg. Auto Points")
                single_team_worksheet.write(25, 0, "Avg. Teleop Points")
                single_team_worksheet.write(26, 0, "Avg. Climb Points")
                single_team_worksheet.write(27, 0, "Matches on Defense")

                # Write the rank from each category to the sheet

                # Avg total match points
                for i, (avg_match_contrib_team_num, avg_match_contrib) in enumerate(all_team_avg_match_contribution):
                    if team_num == avg_match_contrib_team_num:
                        single_team_worksheet.write(23, 1, i + 1)
                # Auto
                for i, (avg_auto_team_num, avg_auto_pts) in enumerate(all_team_avg_auto):
                    if team_num == avg_auto_team_num:
                        single_team_worksheet.write(24, 1, i + 1)
                # Teleop
                for i, (avg_tele_team_num, avg_tele_pts) in enumerate(all_team_avg_tele):
                    if team_num == avg_tele_team_num:
                        single_team_worksheet.write(25, 1, i + 1)
                # Climb
                for i, (avg_climb_team_num, avg_climb_pts) in enumerate(all_team_avg_climb):
                    if team_num == avg_climb_team_num:
                        single_team_worksheet.write(26, 1, i + 1)
                # Defense
                for i, (percent_defense_team_num, percent_defense) in enumerate(all_team_defense_percent):
                    if team_num == percent_defense_team_num:
                        single_team_worksheet.write(27, 1, i + 1)


print("\n> Successfully Created Ouput Workbook\n")
