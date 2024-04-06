'''
    Create a script that will convert mentor's data from mentors.xlsx to mentors.yml file

    TODO:
        Where to get:
            disabled: ?
            sort: ?
            index: ?
            years: ?
'''

import os
import textwrap
import re
import pandas as pd
from ruamel.yaml import YAML
from ruamel.yaml.scalarstring import LiteralScalarString

# Social media
SOCIAL_MEDIA = ['linkedin', 'twitter', 'github', 'medium', 'youtube',\
                'instagram', '//t.', 'meetup', 'slack', 'facebook']
WEBSITE = 'website'
TELEGRAM = 'telegram'

# Indexes used when creating yamal sequences of data
# For mentor's areas block
AREAS_START_INDEX = 9
AREAS_END_INDEX = 13
# For mentor's focus block
FOCUS_START_INDEX = 14
FOCUS_END_INDEX = 18
# For mentor's programming lang. block
PROG_LANG_START_INDEX = 19
PROG_LANG_END_INDEX = 23

TYPE_AD_HOC = "ad hoc"
TYPE_LONG_TERM = "long-term"
TYPE_BOTH = "both"


def get_social_media_links(social_media_links_str):
    '''
        Create social media links displayed on mentor's Skills tab. 
        Links are listed under network field in yaml file.
    '''
    # Split string into list (split on space)
    social_media_links_list = social_media_links_str.split()

    # List of social media links prepared for "network:" block of yaml file.
    # Links are displayed as key-value pairs
    yaml_network_list = []

    for link in social_media_links_list:
        found = 0
        for name in SOCIAL_MEDIA:
            if (link.find(name) != -1):
                # Special case for Telegram
                if (name == "//t."):
                    yaml_network_list.append({TELEGRAM: link})
                else:
                    yaml_network_list.append({name: link})
                found = 1
                break
        if found == 0:
            # Use "webpage" as the key if not one of the known social media links
            yaml_network_list.append({WEBSITE: link})

    # Return mentor's list of social media links
    return yaml_network_list

def get_data_sequence_as_list(mentor_data, start_index, end_index):
    '''
        Mapping Scalars to Sequences using Python lists:
            Mappings use a colon and space (“: ”) to mark each key/value pair.
            Block sequences indicate each entry with a dash and space (“- ”).
    '''
    data_sequence = []
    for entry in range(start_index, end_index+1):
        # Check that the value is not Nan (if Excel cell empty)
        # and remove trailing spaces
        if not pd.isna(mentor_data.iloc[entry]):
            data_sequence.append(mentor_data.iloc[entry].rstrip())

    # Return block sequence of data as a list or en empty string
    if data_sequence:
        return data_sequence
    return ""

def get_numbers_from_string(text_arg, get_max_value=True):
    '''
        Find numbers in a string and convert them to integers.
        Example: If the hours field is of type string (fro example: "5 or more")
        just a digit is extracted.
        text_arg: integer, float or string
        return: 
            - max number or
            - list of numbers or
            - empty str in case of error
    '''
    #TODO: Check - is it possible to have en empty field for available hours?
    if isinstance(text_arg, (int, float)):
        return text_arg
    elif isinstance(text_arg, str):
        digits = [int(num) for num in re.findall(r"\d+", text_arg)]
        if digits:
            if get_max_value:
                return max(digits)
            else:
                return digits
    # Error case
    return ''

def get_multiline_string(long_text_arg):
    '''
        Save strings as yaml multiline strings.
        Use literal block scalar style to keep newlines (using sign |).
    '''
    multiline_str = ''
    if not pd.isna(long_text_arg):
        multiline_str = LiteralScalarString(textwrap.dedent(long_text_arg))
    return multiline_str

def get_mentorship_type(mentorship_type_str):
    '''
        Return mentorship type: ad-hoc, long-term, both or empty str
    '''
    type = mentorship_type_str.lower()

    if TYPE_AD_HOC in type:
        return TYPE_AD_HOC.replace(' ', '-')
    elif TYPE_LONG_TERM in type:
        return TYPE_LONG_TERM
    elif TYPE_BOTH in type:
        return TYPE_BOTH

    return ""

def create_mentors_yml_file(mentors_data):
    '''
        Create mentors.yml file
        Input args:
            mentors_data: list of dictionaries
    '''
    # Relative path
    file_output = ".\\tools\\mentors.yml"

    with open(file_output, 'w', encoding = "utf-8") as output_fp:
        yaml = YAML()
        # TODO: check flow styles
        yaml.default_flow_style = False
        # Set indenting rules in yaml file
        yaml.indent(mapping=2, sequence=4, offset=2)
        # Write data
        yaml.dump(mentors_data, output_fp)
    print(f"File: {file_output} is successfully written.")

def read_mentors_xlsx_file(xlsx_file):
    '''
        Read mentors.xlsx file and prepare data in yaml format
        Input args:
            xlsx_file: Excel file with mentors' data
    '''
    # Mentors' data saved as a list of dictionaries
    mentors = []

    # Load data from the Mentors sheet into a DataFrame object
    df_excel = pd.read_excel(xlsx_file, sheet_name="Mentors")

    for i in range(0, len(df_excel)):
        # Get table row
        print(f"Current index: {i}")
        mentor_row = df_excel.iloc[i]

        # Save strings as yaml multiline strings
        mentee_str = get_multiline_string(mentor_row.iloc[24])
        bio_str = get_multiline_string(mentor_row.iloc[25])
        extra_str = get_multiline_string(mentor_row.iloc[26])

        # Save mentor's areas, focus and programming languages data in a list format
        # List format converts to block sequences in yaml file
        # Block sequences indicate each entry with a dash and space ("- ")
        areas = get_data_sequence_as_list(mentor_row,
                                            AREAS_START_INDEX,
                                            AREAS_END_INDEX)
        focus = get_data_sequence_as_list(mentor_row,
                                            FOCUS_START_INDEX,
                                            FOCUS_END_INDEX)
        programming_languages = get_data_sequence_as_list(mentor_row,
                                                            PROG_LANG_START_INDEX,
                                                            PROG_LANG_END_INDEX)

        type_of_mentorship = get_mentorship_type(mentor_row.iloc[2])
        network_list = get_social_media_links(mentor_row.iloc[27])
        hours_per_month = get_numbers_from_string(mentor_row.iloc[28])
        max_experience = get_numbers_from_string(mentor_row.iloc[8])

        # TODO: Temp solution - change
        disabled = False
        matched = True
        sort = 10
        index = i+1
        image = ""

        mentor = {'name': mentor_row.iloc[0],
                'disabled': disabled,
                'matched': matched,
                'sort': sort,
                'hours': hours_per_month,
                'type': type_of_mentorship,
                'index': index,
                'location': mentor_row.iloc[4],
                'position': f"{mentor_row.iloc[6]}, {mentor_row.iloc[7]}",
                'bio': bio_str,
                'image': image,
                'languages': mentor_row.iloc[5],
                'skills':{
                    'experience': mentor_row.iloc[8],
                    'years': max_experience,
                    'mentee': mentee_str,
                    'areas': areas,
                    'languages': ', '.join(programming_languages),
                    'focus': focus,
                    'extra': extra_str,
                    },
                'network': network_list,
                }
        mentors.append(mentor)
    return mentors

if __name__ == "__main__":
    print(os.getcwd())
    FILE_PATH = ".\\tools\\mentors.xlsx"
    list_of_mentors = read_mentors_xlsx_file(FILE_PATH)
    create_mentors_yml_file(list_of_mentors)
