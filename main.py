import json
import os
import pathlib

import inflection
import pandas as pd


def read_excel(filepath, columns=None):
    df = pd.read_excel(filepath, usecols=columns)
    return df.to_dict()


def create_django_fixture(df_dict, app, model):
    results = []
    first_value_row_length = len(next(iter(df_dict.values())))
    for index in range(1, first_value_row_length + 1):
        record = {
            "model": f"{app}.{model}",
            "pk": index,
            "fields": {}
        }
        for key, value in df_dict.items():
            record["fields"][key] = value[index - 1] if value[index - 1] else None
        results.append(record)
    return results


def create_file(output_name, data):
    # create folders
    pathlib.Path('/'.join(output_name.split('/')[:-1])).mkdir(parents=True, exist_ok=True)
    with open(output_name, "w", encoding='UTF-8') as file:
        file.write(data)


def get_file_fullname(name):
    # https://stackoverflow.com/a/11969014/14506165
    for file in os.listdir('.'):
        if os.path.isfile(file) and name in file:
            return file
    return name


def get_inputs():
    # https://stackoverflow.com/a/17328907/14506165
    app_input = inflection.underscore(input("Enter the app name of the model: "))
    model_input = inflection.underscore(input("Enter the name of the model: "))
    excel_input = get_file_fullname(input("Enter the excel file name: "))
    return app_input, model_input, excel_input


def main():
    app_name, model_name, excel_filename = get_inputs()
    # send columns parameter if you want to filter data by columns
    # example "A,B,C,F"
    df_dict = read_excel(excel_filename)
    fixture_json = create_django_fixture(df_dict, app_name, model_name)
    create_file(f'fixtures/{model_name}.json', json.dumps(fixture_json, ensure_ascii=False))


if __name__ == '__main__':
    main()
