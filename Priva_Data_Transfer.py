
"""The following python programmed was authored, compiled and distributed by Tyler Thompson, ATL. No outside
sources were used or copied in the creation of this file."""

from os import path, listdir
from re import compile, search, split
from time import sleep

import numpy as np
from numpy import nan
from openpyxl import Workbook
from pandas import DataFrame, read_csv, ExcelFile, read_excel, ExcelWriter, to_datetime


class NoValueError(Exception):
    """Available to catch Exceptions where no measured datapoint is available."""

    pass


class NoUsernameFound(Exception):
    """Available to catch circumstances where the logic cannot drill down into a confirmed username."""

    def __init__(self):
        print("""A valid username could not be found. Usernames are identified from the available folders in the 
        C:\\Users\\ directory. If the user name is present, assumptions are that it makes are the username does not 
        have any type of punctuation, they do not have numbers, they do not have standard words, and, when found, 
        they have an accessible directory.

        The intent is for both the program and the downloaded PRIVA .csv files to live on the Desktop. If both of those
        conditions are satisfied, then your desktop path is different than expected; e.g. not C:\\Users\\YourName\\..... 
        Another alternative is that your desktop is not directly linked to your OneDrive, or that your OneDrive desktop 
        uses the username verbatim instead of the "OneDrive - Gotham Greens Holdings, LLC" label.""")

        sleep(20)

        return


class BaseDirectoryError(Exception):
    """Available to catch circumstances where the user has put the files in a directory that is unreachable by
    the software program."""

    def __init__(self):
        print("""The software encountered an error related to the location of the input files downloaded from Priva,
        an issue related to the setup of your desktop, or both. The former requires that both the destination file and 
        the files to be processed are located in a "desktop" folder. The latter requires that desktop folder to be in 
        either a pathway resembling C:\\Users\\YourName\\Desktop or C:\\Users\\OneDrive - Gotham Greens Holdings, LLC.
        If both are satisfied, then this software is not customized for your machine.""")

        sleep(20)

        return


class UsedFileHolder:
    """Class object to keep track of the files that are explored for Priva downloads."""

    list_of_confirmed_files = []

    def __init__(self):
        return

    def test_list(self, file_string):
        """The function takes a possible filename string and stores it so that search and findings are not repeated."""

        if file_string not in self.list_of_confirmed_files:

            self.list_of_confirmed_files.append(file_string)

            return True
        else:
            return False


def has_a_value(func):
    """Wrapper checks whether the data row from the open priva file has a value."""

    def check_for_value(*args):
        """Return the function and its args, or raise an exception"""

        # Priva delimits with a semicolon
        values = split(";", args[1])

        if bool(values[-1]):
            return func(*args)
        else:
            raise NoValueError

    return check_for_value


def onedrive_or_desktop(func):
    """To test whether the pathways to the user's Desktop folder is any of the options evaluated below."""

    def insert_correct_onedrive(*args):

        if args[0] is not None:

            # For a OneDrive linked desktop.
            if path.isdir(f"C:\\Users\\{args[0]}\\OneDrive - Gotham Greens Holdings, LLC\\Desktop"):

                desktop_linkage = r"OneDrive - Gotham Greens Holdings, LLC\Desktop"
                return func(args[0], desktop_linkage)

            # For a straight desktop path.
            elif path.isdir(f"C:\\Users\\{args[0]}\\Desktop"):

                desktop_linkage = r"Desktop"
                return func(args[0], desktop_linkage)

            else:

                raise BaseDirectoryError

    return insert_correct_onedrive


@has_a_value
def assign_a_row(df, confirmed_row, interval, designation, user_label) -> None:
    """Perform the assignment of the downloaded data row to the temporary holding dataframe."""

    df.loc[len(df.index)] = ([str(interval)] + [str(designation)] + [str(user_label)] +
                             split(";", confirmed_row))

    return


@onedrive_or_desktop
def create_data_filename_by_user(found_user=None, str_suffix=None) -> str:
    """Return a modified raw string having the users directory name as part of the downloaded data
    file(s) from Priva."""

    base_str = f"C:\\Users\\***\\{str_suffix}\\+++.csv"

    return base_str.replace('***', found_user)


@onedrive_or_desktop
def create_dest_filename_by_user(found_user=None, str_suffix=None) -> str:
    """Return a modified raw string having the users directory name as part of the destination file."""

    if found_user == "TylerThompson":

        sharepoint_redirect = r"Gotham Greens Holdings, LLC\Growing Work Space - ATL\CLIMATE & IRRIGATION"
        base_str = f"C:\\Users\\***\\{sharepoint_redirect}\\PRIVA_DATA.xlsx"

    elif found_user == "GrahamTucker":

        # Nicks path
        # "C:\Users\GrahamTucker\Gotham Greens Holdings, LLC\Growing Work Space - CHI2\CLIMATE and IRRIGATION\PRIVA_DATA.xlsx"
        sharepoint_redirect = r"Gotham Greens Holdings, LLC\Growing Work Space - CHI2\CLIMATE and IRRIGATION"
        base_str = f"C:\\Users\\***\\{sharepoint_redirect}\\PRIVA_DATA.xlsx"

    elif found_user == "NickBellizzi":

        sharepoint_redirect = r"Gotham Greens Holdings, LLC\Growing Work Space - YIELDS + CLIMATE"
        base_str = f"C:\\Users\\***\\{sharepoint_redirect}\\PRIVA_DATA.xlsx"

    else:

        base_str = f"C:\\Users\\***\\{str_suffix}\\PRIVA_DATA.xlsx"

    return base_str.replace('***', found_user)


def create_base_directory(confirmed_username) -> str:
    """The assumption is that every Gotham employee has the same base route to their employee account."""

    base_str = "C:\\Users\\***\\"

    return base_str.replace('***', confirmed_username)


def confirm_user_found(users_list=None) -> str:
    """The function is updated to look through a users possible. Assumptions that it makes are the username does not
    have any type of punctuation, they do not have numbers, they do not have standard words, and, when found, they have
    an accessible directory."""

    user_name_match = None

    if users_list is not None:

        # These are the literal, std options for non-user names in the C://Users directory.
        # This could go wrong if the users name has a period(.), e.g. Jr., in it
        # Add here if different, and append any valid names to possible_users.
        name_regex = compile(r"\s|\d|[(]|[.]|default|user|all|All|User|Default|Public|desktop|Desktop")
        possible_users = []

        for username in users_list:

            test_name = search(name_regex, username)

            if test_name is None:
                possible_users.append(username)

        if len(possible_users) == 1:
            user_name_match = possible_users[0]
        elif len(possible_users) < 1:
            raise NoUsernameFound
        else:
            names = iter(possible_users)

            while True:
                try:
                    username = next(names)
                    listdir(f"C:\\Users\\{username}\\")
                    user_name_match = username
                    break

                except PermissionError:
                    # Catches instances of Company IT Professionals.
                    pass

                finally:

                    if user_name_match is not None:
                        break

    if user_name_match is not None:
        # User Console Message
        print(f'Username found {user_name_match}')

        return user_name_match

    else:
        raise NoUsernameFound


def report_by_week_generator(found_user=None) -> []:
    """The generator attempts combinations of report name and week number. It only yields when a valid path is found."""

    # Read in the file(s) downloaded from the controller Priva. This is not the destination folder.
    data_file_path_with_user = create_data_filename_by_user(found_user)
    # The list containing all file names in the directory dir_holding_data_files
    dir_holding_data_files = data_file_path_with_user.replace(r'+++.csv', '')
    dir_files = listdir(dir_holding_data_files)

    # Priva does have a week == 53...
    calendar_range = [i for i in range(1, 54, 1)]
    data_periods = ['week', 'period', 'quarter', 'year']

    reports = {'General_Report_Compartment': calendar_range,
               'General_Report_Company': calendar_range,
               'Climate_Report': calendar_range,
               'Lighting_Report': calendar_range,
               'Energy_Report': calendar_range,
               'Water_Report': calendar_range,
               'Valves_Report': calendar_range}

    for iterated_file in dir_files:

        for key in reports.keys():

            for values in reports[key]:

                for period in data_periods:

                    file_name = rf'{key}_regex_{period}_{values}'
                    full_path_plus_expression = file_name.replace('regex', r"(?P<user_tag>([A-Za-z0-9]+))")
                    data_name_regex = compile(full_path_plus_expression)
                    file_name_string = search(data_name_regex, iterated_file)

                    if file_name_string and Explored_Files.test_list(file_string=file_name_string[0]):

                        checked_file_path = data_file_path_with_user.replace('+++', file_name_string[0])

                        # Check to see if the literal options can be found as a file.
                        if path.isfile(checked_file_path):
                            yield [checked_file_path, file_name_string[0], file_name_string.group("user_tag")]
                        else:
                            continue

                    else:
                        continue

    return


def new_data_type_prep(df=None, parent_file_name=None):
    """The function applies the indicated data types to the dataframe columns. Timing of the call occurs before
    the rows are transferred into the memory-bound Excel file. The application of this function is certainly
    redundant to the final file read where deduplication and full data typing occurs. However, the intent is to
    prevent unforeseen exceptions."""

    if df is not None:
        # Data types: Everything is a category unless indicated otherwise.
        df.iloc[:, 0] = df.iloc[:, 0].astype(int)
        df.iloc[:, [1, 2, 3, 4, 5, 6, 7]] = df.iloc[:, [1, 2, 3, 4, 5, 6, 7]].astype("category")
        df.iloc[:, -3] = to_datetime(df.iloc[:, -3], errors='ignore', format="%m/%d/%y").dt.date
        df.iloc[:, -2] = to_datetime(df.iloc[:, -2], errors='ignore', format="%m/%d/%y").dt.date
        df.iloc[:, -1] = df.iloc[:, -1].astype(float)

        return df

    else:
        # A blank DF is returned to keep the programming running.
        # User Console Message
        print(f"An error occurred when transferring file {parent_file_name}.")

        return DataFrame()


if __name__ == "__main__":
    """The program initializing statement."""

    # User Console Message
    print("Priva_Data_Transfer program loaded and initialized.")

    # Class data object holding used files to prevent unneeded replication of files processed.
    Explored_Files = UsedFileHolder()

    user: str = confirm_user_found(users_list=listdir("C:\\Users\\"))
    current_working_dir = create_base_directory(user)

    # Ultimate destination file creation and opening.
    dest_file_path = create_dest_filename_by_user(user)

    # Test whether the file already exists. If not, create the file and assign it.
    if not path.isfile(dest_file_path):

        # User Console Message
        print("Creating the destination folder PRIVA_DATA.xlsx")
        wb = Workbook()
        wb.save(dest_file_path)
        wb.close()

    final_file = ExcelFile(dest_file_path, engine="openpyxl")
    destination_file_initial_read = read_excel(final_file, keep_default_na=False, parse_dates=True,
                                               sheet_name="Sheet", header=0)
    bottom_row = len(destination_file_initial_read)

    # User Console Message
    print("The destination file, PRIVA_DATA, is ready and is reused by this program during future data row additions."
          if bottom_row == 0 else f"File has {bottom_row} prior data rows.")
    del destination_file_initial_read
    final_file.close()

    # Creates a generator for all Priva file possibilities for a given user and save tag.
    possible_files = report_by_week_generator(user)

    # The interval is taken from the downloaded file name.
    interval_number_regex = compile(r"\d+(?=\.csv)")
    interval_type_regex = compile(r"week|period|quarter|year]")

    with ExcelWriter(dest_file_path, mode='a', if_sheet_exists='overlay',
                     engine='openpyxl', engine_kwargs={"data_only": True, "keep_vba": False}) as writer:

        # Valid until filename combinations are exhausted
        while True:

            try:
                # Only valid file paths will make it out of the generator.
                data_file_path, file, user_entry = next(possible_files)
                weekly_file = read_csv(data_file_path, header=0)

            except StopIteration:
                break

            else:
                # A temporary dataframe to hold the collected data.
                new_info_df = DataFrame(columns=None)

                # Week number is not always in Priva data, so it is taken from the filepath.
                interval_number = search(interval_number_regex, data_file_path)[0]
                interval_type = search(interval_type_regex, data_file_path)[0]

                # A generator for all the data rows in the file.
                row_iterator = weekly_file.itertuples()

                # Add the column labels to the first row in the sheet IF the file is new.
                if len(new_info_df.values) == 0:

                    column_header = ["interval"] + ["interval_type"] + ['grower_label'] + split(";", weekly_file.columns[0])

                    for name in column_header:
                        new_info_df[name] = nan

                    new_info_df.reset_index(drop=True)

                # Valid until data rows in the found, downloaded files are exhausted.
                while True:

                    # Must be a string to prevent a TypeError in the wrapper above.
                    row = ''

                    try:
                        # noinspection PyTypeChecker
                        row = next(row_iterator)[1]

                    except StopIteration:
                        break

                    else:
                        try:
                            # This will add the parsed data into the new DF columns.
                            assign_a_row(new_info_df, row, interval_number, interval_type, user_entry)

                        except NoValueError:
                            pass

                count_new_values_found = len(new_info_df)

                # User Console Message
                print(f'Found {count_new_values_found} values from {file}.')

                if count_new_values_found:
                    header = True if 0 == bottom_row else False
                    new_info_df_typed = new_data_type_prep(new_info_df, file)
                    new_info_df_typed.to_excel(writer, sheet_name="Sheet",
                                               startrow=bottom_row, index=False, header=header)
                    bottom_row += count_new_values_found

    # User Console Message
    print(f"Checking for any duplicated rows and applying data types.")

    with ExcelWriter(dest_file_path, mode='a', if_sheet_exists='replace',
                     engine='openpyxl', engine_kwargs={"data_only": True, "keep_vba": False}) as writer:

        final_file = ExcelFile(dest_file_path, engine="openpyxl")
        data_types_by_column = {'interval': np.int32, 'interval_type': "category", 'grower_label': None,
                                'label': "category", 'pcu': "category", 'type_1': "category", 'idx_1': None,
                                'type_2': "category", 'idx_2': None, 'value': np.float64}
        final_file_as_df = read_excel(final_file,  sheet_name="Sheet", header=0,
                                      keep_default_na=True, parse_dates=True, dtype=data_types_by_column)

        # possibly https://www.youtube.com/watch?v=lCFEzRaqoJA
        final_file_as_df.drop_duplicates(keep='last', inplace=True,
                                         subset=['interval', 'interval_type', 'label',
                                                 'type_1', 'idx_1', 'startdate', 'value'])

        # User Console Message
        print(f"Merging {len(final_file_as_df)} new and former rows.")

        final_file_as_df.to_excel(writer, sheet_name="Sheet", startrow=0, index=False,
                                  header=True, float_format="%.3f", engine='openpyxl')

        final_file.close()

# User Console Message
print('Program Completed. You can delete your downloaded Priva file(s).')

# Pauses program to allow the user to read the messages.
sleep(8)
