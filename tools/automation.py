"""
    Create a script that will convert mentor's data from mentors.xlsx to mentors.yml file
"""

import os
import textwrap
import re
import pandas as pd
from ruamel.yaml import YAML
from ruamel.yaml.scalarstring import LiteralScalarString

SOCIAL_MEDIA = ['linkedin', 'twitter', 'github', 'medium', 'youtube',\
                'instagram', '//t.', 'meetup', 'slack', 'facebook']
WEBSITE = 'website'
TELEGRAM = 'telegram'

# Indexes for creating yaml sequences of data
AREAS_START_INDEX = 9
AREAS_END_INDEX = 13
FOCUS_START_INDEX = 14
FOCUS_END_INDEX = 18
PROG_LANG_START_INDEX = 19
PROG_LANG_END_INDEX = 23

TYPE_AD_HOC = "ad hoc"
TYPE_LONG_TERM = "long-term"
TYPE_BOTH = "both"
IMAGE_FILE_PATH = "assets/images/mentors/"
IMAGE_SUFFIX = ".jpeg"


def get_social_media_links(social_media_links_str):
    """
    Prepare mentor's social media links for yaml network sequence.
    """
    network_list = []
    social_media_links_list = social_media_links_str.split()

    for link in social_media_links_list:
        found = 0
        for name in SOCIAL_MEDIA:
            if link.find(name) != -1:
                if name == "//t.":
                    network_list.append({TELEGRAM: link})
                else:
                    network_list.append({name: link})
                found = 1
                break
        if found == 0:
            network_list.append({WEBSITE: link})

    return network_list


def get_yaml_block_sequence(mentor_data, start_index, end_index):
    """
    Yaml block sequence is presented as a list of entries marked with
    dash and space (“- ”). 
    """
    block_sequence_list = []
    for entry in range(start_index, end_index+1):
        if not pd.isna(mentor_data.iloc[entry]):
            block_sequence_list.append(mentor_data.iloc[entry].rstrip())
    return block_sequence_list


def extract_numbers_from_string(text_arg, get_max_value=True):
    """
    Extract numbers and convert them to integers.
    """
    if isinstance(text_arg, (int, float)):
        return text_arg

    if isinstance(text_arg, str):
        digits = [int(num) for num in re.findall(r"\d+", text_arg)]
        if digits:
            if get_max_value:
                return max(digits)
            return digits

    return ""


def get_multiline_string(long_text_arg):
    """
    Save strings as yaml multiline strings.
    Use literal block scalar style to keep newlines (using sign |).
    """
    multiline_str = ''
    if not pd.isna(long_text_arg):
        multiline_str = LiteralScalarString(textwrap.dedent(long_text_arg))

    return multiline_str


def get_mentorship_type(mentorship_type_str):
    """
    Returns ad-hoc, long-term, both or empty str
    """
    mentorship_type = mentorship_type_str.lower()

    if TYPE_AD_HOC in mentorship_type:
        return TYPE_AD_HOC.replace(' ', '-')
    elif TYPE_LONG_TERM in mentorship_type:
        return TYPE_LONG_TERM
    elif TYPE_BOTH in mentorship_type:
        return TYPE_BOTH

    return ""


def write_yml_file(file_path, mentors_data, mode):
    """
    Create new or append to mentors.yml file
    :mentors_data: list of dictionaries
    :mode: 'w' or 'a'
    """
    with open(file_path, mode, encoding = "utf-8") as output_yml:
        yaml = YAML()

        # TODO: check if flow styles def needed
        yaml.default_flow_style = False

        if mode == 'w':
            yaml.indent(mapping=2, sequence=4, offset=2)

        yaml.dump(mentors_data, output_yml)
    print(f"File: {file_path} is successfully written.")


def read_yml_file(file_path):
    """
    Read yml file
    """
    with open(file_path, 'r', encoding="utf-8") as input_yml:
        yaml=YAML(typ='safe')
        yml_dict = yaml.load(input_yml)
        print(f"File: {file_path} is successfully read.")
    return yml_dict


def xlsx_to_yaml_parser(mentor_row, mentor_index):
    """
    Prepare mentor's excel data for yaml format
    """

    mentee_str = get_multiline_string(mentor_row.iloc[24])
    bio_str = get_multiline_string(mentor_row.iloc[25])
    extra_str = get_multiline_string(mentor_row.iloc[26])

    areas = get_yaml_block_sequence(mentor_row,
                                    AREAS_START_INDEX,
                                    AREAS_END_INDEX)
    focus = get_yaml_block_sequence(mentor_row,
                                    FOCUS_START_INDEX,
                                    FOCUS_END_INDEX)
    programming_languages = get_yaml_block_sequence(mentor_row,
                                                    PROG_LANG_START_INDEX,
                                                    PROG_LANG_END_INDEX)

    type_of_mentorship = get_mentorship_type(mentor_row.iloc[2])
    network_list = get_social_media_links(mentor_row.iloc[27])
    hours_per_month = extract_numbers_from_string(mentor_row.iloc[28])
    max_experience = extract_numbers_from_string(mentor_row.iloc[8])

    mentor_disabled = True
    mentor_matched = False
    mentor_sort = 10

    # Left commented since the code might be used in the later version (if decided to
    # add default picture until the mentor's image is not available)
    # mentor_image = os.path.join(IMAGE_FILE_PATH, str(mentor_index) + IMAGE_SUFFIX)

    mentor_image =  f"Download image from: {mentor_row.iloc[33]}"

    mentor = {'name': mentor_row.iloc[0],
            'disabled': mentor_disabled,
            'matched': mentor_matched,
            'sort': mentor_sort,
            'hours': hours_per_month,
            'type': type_of_mentorship,
            'index': mentor_index,
            'location': mentor_row.iloc[4],
            'position': f"{mentor_row.iloc[6]}, {mentor_row.iloc[7]}",
            'bio': bio_str,
            'image': mentor_image,
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
    return mentor


def get_all_mentors_data_in_yml_format(xlsx_file_path):
    """
    Read all mentors from Excel sheet.
    Prepare data for writting to yaml file.
    """
    # list of dict
    mentors = []

    df_mentors = pd.read_excel(xlsx_file_path, sheet_name="Mentors")

    for row in range(0, len(df_mentors)):
        mentor =  xlsx_to_yaml_parser(df_mentors.iloc[row], row+1)
        mentors.append(mentor)
    return mentors


def get_new_mentors_data_in_yml_format(yml_file_path, xlsx_file_path):
    """
    Read just new mentors from Excel sheet
     - start reading xlsx Mentors from the row 88 (from the date 03/04/2024)
     - find diff. between existing yml and xlsx
    Prepare data for writting to yaml file.
    """
    # list of dict
    mentors = []

    # Get mentors' names and indexes from yml file
    mentors_yml_dict = read_yml_file(yml_file_path)
    mentors_names_yml = [sub['name'].lower() for sub in mentors_yml_dict]
    mentors_indexes = [sub['index'] for sub in mentors_yml_dict]

    # Highest index is used as the reference point from which
    # new indexes are calculated
    new_index = max(mentors_indexes) + 1

    df_mentors = pd.read_excel(xlsx_file_path, sheet_name="Mentors", skiprows=86)

    # Get mentors' names from xlsx file
    mentors_names_xlsx = {}
    for i in range(0, len(df_mentors)):
        mentors_names_xlsx[i] = df_mentors.iloc[i].values[0].lower()

    for row, name in mentors_names_xlsx.items():
        if name not in mentors_names_yml:
            mentor = xlsx_to_yaml_parser(df_mentors.iloc[row], new_index)
            new_index += 1
            mentors.append(mentor)
    return mentors


if __name__ == "__main__":
    # TODO: Allow cmd line execution of the script:
    #  line parameters:
    #    path to mentors.xlsx
    #    path to mentors.yml
    #    choices: append or create new

    # While in development work with temp. yml and xlsx files
    # TODO: Change paths when feature complete
    FILE_PATH = ".\\tools\\mentors.xlsx"
    YML_FILE_PATH = ".\\tools\\mentors.yml"

    # APPEND TO EXISTING YML
    list_of_mentors = get_new_mentors_data_in_yml_format(YML_FILE_PATH, FILE_PATH)

    if list_of_mentors:
        write_yml_file(YML_FILE_PATH, list_of_mentors, 'a')

    # CREATE NEW YML
    # TODO: When creating new file, indexes of the mentors
    # in the current yml file must be used in the new yml file.
    # list_of_mentors = get_all_mentors_data_in_yml_format(FILE_PATH)
    # write_yml_file(YML_FILE_PATH, list_of_mentors, 'w')
