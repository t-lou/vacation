#! /bin/env python3
import re
import urllib
import urllib.request
import datetime
import yaml

import pandas


def grab_holiday(year: int, code: str) -> list:
    '''
    Get the list of public holidays from internet.

    Args:
        year: the year to look for.
        code: the region/conutry to check.
    Return:
        List of strings representing the date of public holidays.
    '''
    url = f'https://www.qppstudio.net/publicholidays{year}/{code}.htm'
    pattern = '<time datetime="\d\d\d\d-\d\d-\d\d">'
    text = urllib.request.urlopen(url).read().decode('utf8')
    matches = re.findall(pattern, text)
    return set(it[21:-2] for it in matches)


def load_config() -> dict:
    '''
    Get the configuration, like the year, countries/regions.

    Return:
        Configuration.
    '''
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
        country: grab_holiday(this_year, country)
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

    taken = '  '
    weekend = ' '
    free = ''
    colorization = {
        taken: 'background-color: lightblue',
        weekend: 'background-color: lightgrey',
    }
    col_width = 14

    with pandas.ExcelWriter(config['output'], engine='xlsxwriter') as writer:
        for month, month_data in data.items():
            dates = sorted(list(month_data.keys()))
            to_write = {'Date': dates}
            to_write.update({
                country_display: [
                    taken if country_code in month_data[date] else free
                    for date in dates
                ]
                for country_code, country_display in countries.items()
            })
            to_write['Weekend'] = [
                weekend if 'weekend' in month_data[date] else free
                for date in dates
            ]
            to_write.update({
                person: [free] * len(month_data)
                for person in config['persons']
            })
            pandas.DataFrame(to_write).style.applymap(lambda x: colorization[
                x] if x in colorization else '').to_excel(writer,
                                                          sheet_name=month,
                                                          index=False)
            writer.sheets[month].set_column(1, len(to_write) - 1, col_width)


if __name__ == '__main__':
    main()
