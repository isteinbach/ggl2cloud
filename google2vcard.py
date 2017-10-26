#!/usr/bin/python3

import csv
import itertools
import sys
import vobject
###from collections import defaultdict


def de_star(value):
    is_starred = value.startswith('* ')
    return (value[2:] if is_starred else value, is_starred)

class SimpleMultiEntry(object):
    def __init__(self, row, attribute):
        self.entry = []

        for num in itertools.count(1):
            prefix = attribute + ' ' + str(num) + ' - '
            value_key = prefix + 'Value'

            if value_key not in row:
                break # no more entries

            entry_value = row[value_key]

            if entry_value != '':
                type_key = prefix + 'Type'
                (entry_type, is_starred) = de_star(row[type_key]) if type_key in row else (None, false)
                new_entry = (entry_value, entry_type)

                if is_starred:
                    self.entry.append(new_entry)
                else:
                    self.entry.insert(0, new_entry)

    def primary(self):
        return self.entry[0] if self.entry else (None,None)

    def entries(self):
        return self.entry


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


#defaultdict(<class 'int'>, {'': 126, 'Other': 68, 'Büro': 1, 'Home': 239, 'Benutzerdefiniert': 3, 'alias': 33, 'Work': 60, 'alt': 1})
#defaultdict(<class 'int'>, {'Main': 47, 'Other': 3, 'Persönlich / Fax': 1, 'Persönlich / Mobile': 2, 'Persönlich': 3, 'Work': 102, 'Büro': 2, 'Home': 120, 'Benutzerdefiniert': 4, 'Home Fax': 2, 'Notruf': 1, 'Pager': 4, 'Mobile': 225, 'Work Fax': 25})

def create_from_name(name):
    card = vobject.vCard()
    card.add('fn')
    card.fn.value = name.fn
    if name.prefix or name.given or name.additional or name.family or name.suffix:
        card.add('n')
        card.n.value = vobject.vcard.Name(prefix=name.prefix, given=name.given, additional=name.additional, family=name.family, suffix=name.suffix)
    return card

def add_phones(card, phones):
    for (n,t) in phones.entries():
        p = card.add('tel')
        p.value = n
        if t:
            mapping = { 'Main': 'VOICE', 'Persönlich / Fax': ['HOME','FAX'], 'Persönlich / Mobile': ['HOME','CELL'], 'Persönlich': 'VOICE', 'Work': ['WORK','VOICE'], 'Büro': 'WORK', 'Home': ['HOME','VOICE'], 'Home Fax': ['HOME','FAX'], 'Pager': 'PAGER', 'Mobile': 'CELL', 'Work Fax': ['WORK','FAX'] }
            p.type_param = mapping[t] if t in mapping else 'OTHER'

def add_emails(card, emails):
    for (a,t) in emails.entries():
        e = card.add('email')
        e.value = a
        if t:
            mapping = { 'Büro': 'WORK', 'Home': 'HOME', 'Work': 'WORK' }
            e.type_param = mapping[t] if t in mapping else 'OTHER'

def add_simple(card, row, column, tag):
    if column in row:
        value = row[column]
        if value:
           card.add(tag).value = value

def create_card(name, emails, phones, row):
    card = create_from_name(name)
    add_phones(card, phones)
    add_emails(card, emails)
    add_simple(card, row, 'Birthday', 'bday')
    add_simple(card, row, 'Notes', 'note')

    return card


def convert(filename):
    with open(filename, 'r', encoding='utf16') as source:
        reader = csv.DictReader(source, delimiter=',', quotechar='"', doublequote=True)
        content = ''
###        email_types = defaultdict(int)
        for row in reader:
            emails = SimpleMultiEntry(row, 'E-mail')
            phones = SimpleMultiEntry(row, 'Phone')
            name = Name(row, emails, phones)
            websites = SimpleMultiEntry(row, 'Website')

            card = create_card(name, emails, phones, row)
            content += card.serialize()
        print(content, end='')


if __name__ == '__main__':
    convert(sys.argv[1])
