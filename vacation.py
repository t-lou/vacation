#! /bin/env python3
import re
import urllib
import urllib.request
import datetime
import yaml

import pandas


def grab_holiday(url: str) -> list:
    pattern = '<time datetime="\d\d\d\d-\d\d-\d\d">'
    text = urllib.request.urlopen(url).read().decode('utf8')
    matches = re.findall(pattern, text)
    return set(it[21:-2] for it in matches)


def build_url(year: int, code: str) -> str:
    return f'https://www.qppstudio.net/publicholidays{year}/{code}.htm'


def load_config() -> dict:
    expected = {'countries': list, 'persons': list, 'year': int, 'output': str}
    config = yaml.safe_load(open('config.yml'))
    assert all(it in config and type(config[it]) == expected[it]
               for it in expected)
    return config


def main():
    config = load_config()
    this_year = config['year']
    countries = {
        country['code']: country['display']
        for country in config['countries']
    }

    holidays = {
        country: grab_holiday(build_url(this_year, country))
        for country in countries.keys()
    }

    data = dict()

    date = datetime.datetime(year=this_year, month=1, day=1)
    one_day = datetime.timedelta(days=1)

    this_month = 1
    while date.year == this_year:
        str_month = date.strftime("%Y-%m")
        str_day = date.strftime("%m-%d")

        if str_month not in data:
            data[str_month] = dict()

        data[str_month][str_day] = [
            country for country in countries if str_day in holidays[country]
        ]

        if date.weekday() >= 5:
            data[str_month][str_day].append('weekend')

        date += one_day

    # https://stackoverflow.com/questions/36694313/pandas-xlsxwriter-format-header
    pandas.io.formats.excel.ExcelFormatter.header_style = None

    with pandas.ExcelWriter(config['output'], engine='xlsxwriter') as writer:
        for month, month_data in data.items():
            dates = sorted(list(month_data.keys()))
            to_write = {'Date': dates}
            to_write.update({
                country_display: [
                    ' ' if country_code in month_data[date] else ''
                    for date in dates
                ]
                for country_code, country_display in countries.items()
            })
            to_write['Weekend'] = [
                ' ' if 'weekend' in month_data[date] else '' for date in dates
            ]
            to_write.update({
                person: [''] * len(month_data)
                for person in config['persons']
            })
            pandas.DataFrame(to_write).style.applymap(
                lambda x: 'background-color: grey'
                if x == ' ' else '').to_excel(writer,
                                              sheet_name=month,
                                              index=False)


main()
