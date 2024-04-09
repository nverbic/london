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

    return ""


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


def write_mentors_yml_file(mentors_data):
    """
    Create mentors.yml file
    :mentors_data: list of dictionaries
    """
    file_output = ".\\tools\\mentors.yml"

    with open(file_output, 'w', encoding = "utf-8") as output_fp:
        yaml = YAML()
        # TODO: check if flow styles def needed
        yaml.default_flow_style = False
        # Indenting rules in yaml file
        yaml.indent(mapping=2, sequence=4, offset=2)
        yaml.dump(mentors_data, output_fp)
    print(f"File: {file_output} is successfully written.")


def xlsx_to_yaml_parser(xlsx_file):
    """
    Prepare mentor's excel file data for yaml format
    """
    mentors = []

    df_excel = pd.read_excel(xlsx_file, sheet_name="Mentors")

    for i in range(0, len(df_excel)):
        mentor_row = df_excel.iloc[i]

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

        mentor_disabled = False
        mentor_matched = False
        mentor_sort = 10

        # TODO: Implement dictionary with email: index pairs,
        # in order to preserve existing indexing
        mentor_index = i+1

        # Commented until metor_index is implemeted
        # mentor_image = os.path.join(IMAGE_FILE_PATH, str(mentor_index) + IMAGE_SUFFIX)
        mentor_image =  ""

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
        mentors.append(mentor)
    return mentors

if __name__ == "__main__":
    print(os.getcwd())
    FILE_PATH = ".\\tools\\mentors.xlsx"
    list_of_mentors = xlsx_to_yaml_parser(FILE_PATH)
    write_mentors_yml_file(list_of_mentors)
