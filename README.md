# Vacation calendar for international teams

Working in an international team may not be so easy, when colleagues in different countries have different national or regional holidays.
Then we don't know when to avoid demos and meetings.
This Python tools aims at creating an Excel file to show when each region has public holiday.

Colleagues can also modify the personal leaves.

## Dependency:
- Python3
- pyyaml
- pandas
- xlsxwriter
- jinja2

## Use

- add countries and regions in config
- add or remove the names for personal availability
- run the script

The content of [qppstudio](https://www.qppstudio.net/worldwide-public-holidays/country-portal.htm) is grabbed.
You may select the country and check, which code to use for a country or region.