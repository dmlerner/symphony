import glob
import os
import xlrd
import xlwt
import collections
import re


def to_sheet(filename):
    return xlrd.open_workbook(filename).sheet_by_index(0)


def parse(filename):
    sheet = to_sheet(filename)
    headers = [h.lower() for h in sheet.row_values(0)]
    parsed = []
    for i in range(1, sheet.nrows):
        parsed.append(dict(zip(headers, sheet.row_values(i))))
    return parsed


def unabbreviate_state(abbreviation):
    abbreviations = {
        'AL': 'Alabama',
        'AK': 'Alaska',
        'AZ': 'Arizona',
        'AR': 'Arkansas',
        'CA': 'California',
        'CO': 'Colorado',
        'CT': 'Connecticut',
        'DE': 'Delaware',
        'FL': 'Florida',
        'GA': 'Georgia',
        'HI': 'Hawaii',
        'ID': 'Idaho',
        'IL': 'Illinois',
        'IN': 'Indiana',
        'IA': 'Iowa',
        'KS': 'Kansas',
        'KY': 'Kentucky',
        'LA': 'Louisiana',
        'ME': 'Maine',
        'MD': 'Maryland',
        'MA': 'Massachusetts',
        'MI': 'Michigan',
        'MN': 'Minnesota',
        'MS': 'Mississippi',
        'MO': 'Missouri',
        'MT': 'Montana',
        'NE': 'Nebraska',
        'NV': 'Nevada',
        'NH': 'New Hampshire',
        'NJ': 'New Jersey',
        'NM': 'New Mexico',
        'NY': 'New York',
        'NC': 'North Carolina',
        'ND': 'North Dakota',
        'OH': 'Ohio',
        'OK': 'Oklahoma',
        'OR': 'Oregon',
        'PA': 'Pennsylvania',
        'RI': 'Rhode Island',
        'SC': 'South Carolina',
        'SD': 'South Dakota',
        'TN': 'Tennessee',
        'TX': 'Texas',
        'UT': 'Utah',
        'VT': 'Vermont',
        'VA': 'Virginia',
        'WA': 'Washington',
        'WV': 'West Virginia',
        'WI': 'Wisconsin',
        'WY': 'Wyoming',
        'DC': 'District of Columbia'
    }
    return abbreviations.get(abbreviation.upper(), abbreviation)


def to_converter(filename, key, key_type=str):
    parsed = parse(filename)
    conversions = {}
    for row in parsed:
        conversions[key_type(row[key])] = row
    return conversions


def to_location(facilitator):
    zip_code = facilitator['zip code']
    if re.match(r'\d+', zip_code):
        zip_code = int(zip_code)
        if zip_code in pa_zipcodes:
            return pa_zipcodes[zip_code]['county']
    if 'posttown' in facilitator:
        location = unabbreviate_state(facilitator['state']), facilitator['posttown']
    else:
        location = unabbreviate_state(facilitator['state / province code']), facilitator['city']
    if location == ('', ''):
        return None
    return location


def remove_suffix(filename):
    return filename[:filename.rindex('.')]


class Concert:
    concerts = []

    def __init__(self, filename):
        self.filename = remove_suffix(filename)
        self.facilitators = parse(filename)
        Concert.concerts.append(self)

    def count_attendees(self):
        return sum(int(f['count']) for f in self.facilitators)

    def count_attendees_per_location(self):
        attendees_per_location = collections.Counter()
        for f in self.facilitators:
            location = to_location(f)
            attendees_per_location[location] += int(f['count'])
        return attendees_per_location

    def get_locations(self):
        return set(map(to_location, self.facilitators))

    def get_per(f):
        return {c.filename: f(c) for c in Concert.concerts}

    def get_total(f):
        total = f(Concert.concerts[0])
        for c in Concert.concerts[1:]:
            total += f(c)
        return total

    def get_total_set(f):
        return set.union(*map(f, Concert.concerts))


def write_attendees_per_concert(sheet, start_column):
    attendees_per_concert = Concert.get_per(Concert.count_attendees)
    print(attendees_per_concert, '\n')
    sheet.row(0).write(start_column, 'Concert')
    sheet.row(0).write(start_column + 1, 'Attendance')
    for i, (concert, attendees) in enumerate(attendees_per_concert.items()):
        sheet.row(i + 1).write(start_column, concert)
        sheet.row(i + 1).write(start_column + 1, attendees)
    total_attendees = Concert.get_total(Concert.count_attendees)
    print(total_attendees, '\n')
    sheet.row(i + 2).write(start_column, 'Total')
    sheet.row(i + 2).write(start_column + 1, total_attendees)


def value_sort(d):
    return sorted(list(d.items()), key=lambda x: list(reversed(x)), reverse=True)


def split_counties(x):
    counties = {c for c in x if isinstance(c, str)}
    not_counties = set(x.keys()) - counties
    return (value_sort({c: x[c] for c in C}) for C in (counties, not_counties))


def write_attendees_per_location(sheet, start_column):
    attendees_per_location = Concert.get_total(Concert.count_attendees_per_location)
    print(attendees_per_location, '\n')
    counties, not_counties = split_counties(attendees_per_location)

    sheet.row(0).write(start_column, 'County')
    sheet.row(0).write(start_column + 1, 'Attendance')
    for i, (c, a) in enumerate(counties):
        sheet.row(i + 1).write(start_column, c)
        sheet.row(i + 1).write(start_column + 1, a)

    sheet.row(0).write(start_column + 3, 'State')
    sheet.row(0).write(start_column + 4, 'City')
    sheet.row(0).write(start_column + 5, 'Attendance')
    for i, (k, a) in enumerate(not_counties):
        if k:
            state, city = k
            sheet.row(i + 1).write(start_column + 3, state)
            sheet.row(i + 1).write(start_column + 4, city)
        sheet.row(i + 1).write(start_column + 5, a)


def write_output(filename):
    book = xlwt.Workbook()
    sheet = book.add_sheet(remove_suffix(filename))
    write_attendees_per_concert(sheet, 0)
    write_attendees_per_location(sheet, 3)
    book.save(filename)


def main():
    global pa_zipcodes
    filenames = glob.glob('*.xlsx')

    zipcode_filename = [f for f in filenames if 'Zip' in f][0]
    pa_zipcodes = to_converter(zipcode_filename, 'zip code', int)

    output_filename = "Concert Attendance.xlsx"
    if os.path.exists(output_filename):
        os.remove(output_filename)

    concert_filenames = [f for f in filenames if f != zipcode_filename and f != output_filename]

    for f in concert_filenames:
        Concert(f)
    write_output(output_filename)


if __name__ == '__main__':
    main()
