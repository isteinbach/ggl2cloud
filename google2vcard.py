#!/usr/bin/python3

import csv
import itertools
import sys
import vobject
from collections import defaultdict


def de_multi(value):
    return value.split(' ::: ')

def de_star(value):
    is_starred = value.startswith('* ')
    return (value[2:] if is_starred else value, is_starred)


def remove_key(container, key):
    is_present = key in container
    value = None
    if is_present:
        value = container[key]
        del container[key]
    return value


class SimpleMultiEntry(object):
    def __init__(self, row, attribute, type_mapper):
        self.__attribute = attribute
        self.__entry = []

        for num in itertools.count(1):
            prefix = attribute + ' ' + str(num) + ' - '
            value_key = prefix + 'Value'

            combined_values = remove_key(row, value_key)
            if combined_values is None:
                break # no more entries

            if combined_values:
                type_key = prefix + 'Type'
                type_value = remove_key(row, type_key)
                (original_type, is_starred) = de_star(type_value) if type_value is not None else (None, false)
                entry_type = type_mapper.map(original_type)

                for entry_value in reversed(de_multi(combined_values)):
                    new_entry = (entry_value, entry_type)

                    if is_starred:
                        self.__entry.append(new_entry)
                    else:
                        self.__entry.insert(0, new_entry)

    def primary(self):
        return self.__entry[0] if self.__entry else (None,None)

    def entries(self):
        return self.__entry


class AddressEntry(object):
    def __init__(self, row, prefix, type_mapper):
        self.type_param = type_mapper.map(remove_key(row, prefix + 'Type'))
        self.box = remove_key(row, prefix + 'PO Box')
        self.extended = remove_key(row, prefix + 'Extended Address')
        self.street = remove_key(row, prefix + 'Street')
        self.city = remove_key(row, prefix + 'City')
        self.region = remove_key(row, prefix + 'Region')
        self.code = remove_key(row, prefix + 'Postal Code')
        self.country = remove_key(row, prefix + 'Country')
        self.has_address = self.box or self.extended or self.street or self.city or self.region or self.code or self.country
        self.formatted = remove_key(row, prefix + 'Formatted')

class AddressMultiEntry(object):
    def __init__(self, row, type_mapper):
        self.__entry = []

        for num in itertools.count(1):
            prefix = 'Address ' + str(num) + ' - '

            if not any(key.startswith(prefix) for key in row):
                break # no more entries

            new_entry = AddressEntry(row, prefix, type_mapper)
            self.__entry.append(new_entry)

    def entries(self):
        return self.__entry


class Name(object):
    def __init__(self, row, emails, phones):
        self.fn = remove_key(row, 'Name')
        self.name = self.fn
        self.prefix = remove_key(row, 'Name Prefix')
        self.given = remove_key(row, 'Given Name')
        self.additional = remove_key(row, 'Additional Name')
        self.family = remove_key(row, 'Family Name')
        self.suffix = remove_key(row, 'Name Suffix')

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
        value = remove_key(row, column)
        if value:
            card.add(tag).value = value

def add_multi(card, row, column, tag):
    if column in row:
        multi_values = remove_key(row, column)
        if multi_values:
            # not sure if we should also de_star()...
            card.add(tag).value = de_multi(multi_values)

def add_addresses(card, addresses):
    for entry in addresses.entries():
        if entry.has_address:
            a = card.add('adr')
            a.type_param = entry.type_param
            a.value.box = entry.box
            a.value.extended = entry.extended
            a.value.street = entry.street
            a.value.city = entry.city
            a.value.region = entry.region
            a.value.code = entry.code
            a.value.country = entry.country
        if entry.formatted:
            a = card.add('label')
            a.type_param = entry.type_param
            a.value = entry.formatted


def create_card(name, emails, phones, addresses, websites, row):
    card = create_from_name(name)
    add_attribute_type(card, phones, 'tel')
    add_attribute_type(card, emails, 'email')
    add_addresses(card, addresses)
    add_attribute_type(card, websites, 'url')
    add_multi(card, row, 'Group Membership', 'categories')
    add_simple(card, row, 'Birthday', 'bday')
    add_simple(card, row, 'Nickname', 'nickname')
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
        unhandled = set()
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

        addresses_mapper = Mapper('address types', {
            'Büro': 'WORK',
            'Home': 'HOME',
            'Persönlich': 'HOME',
            'Work': 'WORK'
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
            addresses = AddressMultiEntry(row, addresses_mapper)
            websites = SimpleMultiEntry(row, 'Website', websites_mapper)

            card = create_card(name, emails, phones, addresses, websites, row)
            content += card.serialize()
            for k in row:
                if row[k]:
                    unhandled.add(k)

        print(content, end='')
        emails_mapper.print_unknown()
        phones_mapper.print_unknown()
        addresses_mapper.print_unknown()
        websites_mapper.print_unknown()
        print('Unhandled attributes:', unhandled, file=sys.stderr)


if __name__ == '__main__':
    convert(sys.argv[1])
