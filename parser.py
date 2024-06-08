import pandas as pd
import re

from collections import defaultdict
import json

IGNORE_WORDS = ["de", '']
types = defaultdict(lambda: str)


def get_technological_unit_code(technological_unit: str):
    words = technological_unit.split(" ")
    words = [word for word in words if word not in IGNORE_WORDS]
    return "".join([word[0].upper() if i == 0 else word[0] for i, word in enumerate(words)])


def get_technical_element_class_code(technical_element_class: str):
    words = technical_element_class.split(" ")
    words = [word for word in words if word not in IGNORE_WORDS]
    return "".join([word[0].upper() + word[1] if i == 0 else word[0].upper() for i, word in enumerate(words)])


def normalize_data(text: pd.Series):
    if pd.isna(text).any():
        return text
    else:
        # print((text.str.lower()
        #         .str.strip()
        #         .str.normalize('NFKD')
        #         .str.encode('ascii', errors='ignore')
        #         .str.decode('utf-8')), text)
        return (text.str.lower()
                .str.strip()
                .str.normalize('NFKD')
                .str.encode('ascii', errors='ignore')
                .str.decode('utf-8'))


def retrieve_range(range_boundaries):
    ranges = re.findall('[A-Z]+', range_boundaries)
    rows = [int(number) for number in re.findall('\d+', range_boundaries)]
    df = pd.read_excel(io='source.xlsx', sheet_name='Elementos de observación', skiprows=(rows[0] - 2),
                       nrows=(rows[1] - rows[0] + 1), usecols=(ranges[0] + ":" + ranges[1]), dtype=types)
    df = df.ffill(axis=0)
    df = df.apply(normalize_data)
    return df


def main(ranges):
    df = None
    for each_range in ranges:
        if df is None:
            df = retrieve_range(each_range)
        else:
            temp_df = retrieve_range(each_range)
            df = pd.concat([df, temp_df], axis=1)
            del temp_df

    # TECHNICAL UNITS PARSING
    technological_units = df.iloc[:, 4]
    technological_units_set = set(technological_units.to_numpy())
    technological_unit_codes = [get_technological_unit_code(technological_unit) for technological_unit
                                in technological_units_set]
    technological_unit_dict = dict(zip(technological_units_set, technological_unit_codes))

    # TECHNICAL ELEMENTS PARSING
    technical_element_categories = df.iloc[:, 5]
    technical_element_categories_set = set(technical_element_categories)
    technical_element_codes = [get_technical_element_class_code(technical_element_category)
                               for technical_element_category in technical_element_categories_set]
    technical_element_dict = dict(zip(technical_element_categories_set, technical_element_codes))

    elements = []

    for element in df.to_numpy():
        element[7] = element[7].upper()
        split_code = element[7].split("_")

        location = {
            "floor": split_code[0][0],
            "block": split_code[0][-1],
        }

        number = re.findall('\d+', split_code[2])
        if len(number) > 0:
            location["number"]: int(number[0])

        if len(split_code) > 3:
            location["cardinalPoint"] = split_code[-1]

        elements.append({
            "name": element[5],
            "code": element[7],
            "technicalUnitCode": technological_unit_dict[element[4]],
            "location": location
        })

    with open("output.json", "w") as writer:
        writer.write(json.dumps(elements, indent=4))


if __name__ == '__main__':
    main(["A3:J93", "L3:U93"])
