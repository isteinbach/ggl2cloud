#!/usr/bin/python3

import csv
import itertools
import sys
import vobject
from collections import defaultdict


def de_star(value):
    is_starred = value.startswith('* ')
    return (value[2:] if is_starred else value, is_starred)



class SimpleMultiEntry(object):
    def __init__(self, row, attribute, type_mapper):
        self.__attribute = attribute
        self.__entry = []

        for num in itertools.count(1):
            prefix = attribute + ' ' + str(num) + ' - '
            value_key = prefix + 'Value'

            if value_key not in row:
                break # no more entries

            entry_value = row[value_key]

            if entry_value != '':
                type_key = prefix + 'Type'
                (entry_type, is_starred) = de_star(row[type_key]) if type_key in row else (None, false)
                new_entry = (entry_value, type_mapper.map(entry_type))

                if is_starred:
                    self.__entry.append(new_entry)
                else:
                    self.__entry.insert(0, new_entry)

    def primary(self):
        return self.__entry[0] if self.__entry else (None,None)

    def entries(self):
        return self.__entry


class Name(object):
    def __init__(self, row, emails, phones):
        self.fn = row['Name']
        self.name = row['Name']
        self.prefix = row['Name Prefix']
        self.given = row['Given Name']
        self.additional = row['Additional Name']
        self.family = row['Family Name']
        self.suffix = row['Name Suffix']

        if self.fn == '':
            self.fn = ' '.join(filter(None, [self.prefix, self.given, self.additional, self.family, self.suffix]))

        if self.fn == '':
            (primary_email, __) = emails.primary()
            (primary_phone, __) = phones.primary()
            self.fn = primary_email if primary_email is not None else primary_phone



def create_from_name(name):
    card = vobject.vCard()
    card.add('fn')
    card.fn.value = name.fn
    if name.prefix or name.given or name.additional or name.family or name.suffix:
        card.add('n')
        card.n.value = vobject.vcard.Name(prefix=name.prefix, given=name.given, additional=name.additional, family=name.family, suffix=name.suffix)
    return card


def add_attribute_type(card, collection, attribute):
    for (n,t) in collection.entries():
        a = card.add(attribute)
        a.value = n
        if t:
            a.type_param = t

def add_simple(card, row, column, tag):
    if column in row:
        value = row[column]
        if value:
           card.add(tag).value = value

def create_card(name, emails, phones, websites, row):
    card = create_from_name(name)
    add_attribute_type(card, phones, 'tel')
    add_attribute_type(card, emails, 'email')
    add_attribute_type(card, websites, 'url')
    add_simple(card, row, 'Birthday', 'bday')
    add_simple(card, row, 'Notes', 'note')

    return card


class Mapper(object):
    def __init__(self, attribute, mapping):
        self.__mapping = mapping
        self.__attribute = attribute
        self.__unknown = defaultdict(int)

    def map(self, original):
        is_known = original in self.__mapping
        mapped = self.__mapping[original] if is_known else original
        if original and not is_known:
            self.__unknown[original] += 1
        return mapped

    def print_unknown(self):
        if 0 < len(self.__unknown):
            print('Unknown', self.__attribute+':', self.__unknown, file=sys.stderr)



def convert(filename):
    with open(filename, 'r', encoding='utf16') as source:
        reader = csv.DictReader(source, delimiter=',', quotechar='"', doublequote=True)
        content = ''
#defaultdict(<class 'int'>, {'': 126, 'Other': 68, 'Büro': 1, 'Home': 239, 'Benutzerdefiniert': 3, 'alias': 33, 'Work': 60, 'alt': 1})
        emails_mapper = Mapper('email types', {
            'Büro': 'WORK',
            'Home': 'HOME',
            'Other': 'OTHER',
            'Work': 'WORK'
        })

        phones_mapper = Mapper('phone types', {
            'Büro': 'WORK',
            'Home Fax': ['HOME','FAX'],
            'Home': ['HOME','VOICE'],
            'Main': 'VOICE',
            'Mobile': 'CELL',
            'Other': 'OTHER',
            'Pager': 'PAGER',
            'Persönlich / Fax': ['HOME','FAX'],
            'Persönlich / Mobile': ['HOME','CELL'],
            'Persönlich': 'VOICE',
            'Work Fax': ['WORK','FAX'],
            'Work': ['WORK','VOICE']
        })

        websites_mapper = Mapper('website types', {
            'Büro': 'WORK',
            'Home': 'HOME',
            'Other': 'OTHER',
            'Work': 'WORK'
        })

        for row in reader:
            emails = SimpleMultiEntry(row, 'E-mail', emails_mapper)
            phones = SimpleMultiEntry(row, 'Phone', phones_mapper)
            name = Name(row, emails, phones)
            websites = SimpleMultiEntry(row, 'Website', websites_mapper)

            card = create_card(name, emails, phones, websites, row)
            content += card.serialize()

        print(content, end='')
        emails_mapper.print_unknown()
        phones_mapper.print_unknown()
        websites_mapper.print_unknown()


if __name__ == '__main__':
    convert(sys.argv[1])
