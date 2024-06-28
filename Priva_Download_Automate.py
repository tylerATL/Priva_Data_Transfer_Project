
import pyautogui
import datetime

pyautogui.PAUSE = 2.5


def file_renaming(report,compartment,period,current_week):

    updated_name = f'{report}_{compartment}_{period}_{current_week}'
    return updated_name






# Locations
username_input = (1255,450,0.25)
password_input = (1255,535,0.25)
login_button = (1255,675,0.25)
overviews_menu = (50,700,15)
reports_menu_offset = (100,25,0.5)
general_compartment_report = (200,0,0.5)
energy_report = (0,200,0.5)
valves_report = (0,0,0.5)
period_dropdown = (325,350)
week_dropdown = (565,350)
year_dropdown = (725,350)
compartment_dropdown = (900,350)
calculate_button = (1175,330)
export_button = (1375,330)
priva_save_button = (975,575)
file_name_input = (1275,825)
file_save_button = (1300,975)
close_button = (1125,645)
dropdown_offset = 25



# # Login into Priva
# pyautogui.moveTo(*username_input)
# pyautogui.click(username_input[0],username_input[1],1,_pause=False)
# pyautogui.typewrite('GROWER')
#
# pyautogui.moveTo(*password_input)
# pyautogui.click(password_input[0],password_input[1],1,_pause=False)
# pyautogui.typewrite('grower')
#
# pyautogui.moveTo(*login_button)
# pyautogui.click(login_button[0],login_button[1],1,_pause=False)
#
#
#


# current_week = datetime.date.today().isocalendar()[1]
#
# compartments_in_dropdown = {'General_Report_Compartment':[general_compartment_report, [(101, 2), (201, 1), (301, 1), (302, 1)],['week']],
#                             'Energy_Report':[general_compartment_report, [(101, 2), (201, 1), (301, 1), (302, 1)],['week']]}
#
# for report, params in compartments_in_dropdown.items():
#
#     screen_locations, compartment_settings, periods = params
#
#     for period in periods:
#
#         # Navigate to the reports
#         pyautogui.moveTo(*overviews_menu)
#         pyautogui.moveRel(*reports_menu_offset)
#         pyautogui.moveRel(*screen_locations)
#         pyautogui.click(clicks=1,_pause=False)
#
#
#         for compartment, offset_factor in compartment_settings:
#
#             updated_file_name = file_renaming(report,compartment,period,current_week)
#
#             # Change to the compartments of interest in the dropdown.
#             pyautogui.moveTo(*compartment_dropdown)
#             pyautogui.click(clicks=1,_pause=False)
#             pyautogui.moveRel(0,offset_factor*dropdown_offset)
#             pyautogui.click(clicks=1,_pause=False)
#
#             pyautogui.moveTo(*calculate_button)
#             pyautogui.click(clicks=1,_pause=False)
#
#             # Export and Save the Report
#             pyautogui.moveTo(*export_button)
#             pyautogui.click(clicks=1,_pause=False)
#             pyautogui.moveTo(*priva_save_button)
#             pyautogui.click(clicks=1,_pause=False)
#             pyautogui.moveTo(*file_name_input)
#             pyautogui.typewrite(updated_file_name)
#             pyautogui.moveTo(*file_save_button)
#             pyautogui.click(clicks=1,_pause=False)
#             pyautogui.moveTo(*close_button)
#             pyautogui.click(clicks=1,_pause=False)

pyautogui.moveTo(*overviews_menu)
pyautogui.moveRel(*reports_menu_offset)
pyautogui.moveRel(*general_compartment_report)
pyautogui.moveRel(*energy_report)