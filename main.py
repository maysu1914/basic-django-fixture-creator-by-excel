import json
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


def get_inputs():
    # https://stackoverflow.com/a/17328907/14506165
    app_input = inflection.underscore(input("Enter the app name of the model: "))
    model_input = inflection.underscore(input("Enter the name of the model: "))
    return app_input, model_input


def main():
    app_name, model_name = get_inputs()
    # send columns parameter if you want to filter data by columns
    # example "A,B,C,F"
    df_dict = read_excel('example_data.xlsx')
    fixture_json = create_django_fixture(df_dict, app_name, model_name)
    create_file('fixtures/my_model.json', json.dumps(fixture_json, ensure_ascii=False))


if __name__ == '__main__':
    main()
