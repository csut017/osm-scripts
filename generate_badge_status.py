'''
This script generates an Excel spreadsheet of where the division is for each requirement.
'''

import os
import sys

from datetime import date
import xlsxwriter
from osm import AwardScheme, Connection, Manager


class ReportGenerator(object):

    def __init__(self):
        self._conn = None
        self._mgr = None
        self._section = None
        self._term = None

    def _connect(self):
        self._conn = Connection('secret.json')
        self._conn.connect()

    def _initialise(self):
        self._mgr = Manager()
        self._mgr.load(self._conn)
        if not os.path.exists('temp_images'):
            os.makedirs('temp_images')

    def run(self):
        print('Connecting to OSM...')
        self._connect()
        self._initialise()

        if len(sys.argv) < 3:
            print('ERROR: term and section have not been set! ')
            return

        self._set_term(sys.argv[1:3])
        if self._term is None:
            return

        print('Retrieving badge data...')
        scheme = AwardScheme(self._section.name + '-award.json')
        print('-> Loaded award scheme definition')
        self._term.load_badges(self._conn)
        badge_map = {}
        for badge in self._term.badges:
            badge_map[badge.badge_id] = badge
        print('-> Loaded badges')

        print('Retrieving members...')
        members = self._term.load_members(self._conn)
        print('-> Loaded members')

        print('Generating report...')
        filename = ensureExtension(sys.argv[2]+'-Badge Status', '.xlsx')
        workbook = xlsxwriter.Workbook(filename)
        self._generate_report(scheme, members, badge_map, workbook)

        print('Saving to %s...' % (filename, ))
        workbook.close()

        print('Done')

    def _set_term(self, args):
        term_name = args[0]
        self._set_section(args[1:])

        print('Setting term...')
        if term_name == 'current':
            term = self._section.current_term()
            if term is None:
                print('-> Currently not in a term')
                return
            else:
                self._term = term
                print('-> Term set to %s' % (str(term), ))
                return
        
        for term in self._section.terms:
            if term.name == term_name:
                self._term = term
                print('-> Term set to %s' % (str(term), ))
                return
            
        print('-> Unknown term: %s' % (term_name, ))

    def _set_section(self, args):
        print('Setting section...')
        section = self._mgr.find_section(args[0])
        if section is None:
            print('-> Unknown section: %s' % (args[0], ))
        else:
            self._section = section
            print('-> Section set to %s' % (str(section), ))
    
    def _generate_report(self, scheme, members, badge_map, workbook):
        bold_format = workbook.add_format({'bold': True, 'font_size': 12})
        progress_format = workbook.add_format({'num_format': '0.00'})
        for badge in scheme.badges:
            print('-> Processing ' + badge.name)
            ws_name = badge.name
            if len(ws_name) > 30:
                ws_name = ws_name[0:27] + '...'
            worksheet = workbook.add_worksheet(ws_name)
            worksheet.write('A1', badge.name, bold_format)
            worksheet.write('A2', 'First Name', bold_format)
            worksheet.write('B2', 'Family Name', bold_format)
            worksheet.write('C2', 'Awarded', bold_format)

            row = 3 if badge.group else 2
            member_map = {}
            for member in members:
                if member.patrol != 'Leaders':
                    member_map[member.member_id] = row
                    worksheet.write(row, 0, member.first_name)
                    worksheet.write(row, 1, member.last_name)
                    row += 1

            if not badge.complete_id is None:
                complete_badge = badge_map[badge.complete_id]
                complete_badge.load_progress(self._conn)
                print('-> Loaded "%s"...' % (complete_badge.name,))
                for member in complete_badge.progress:
                    worksheet.write(member_map[member.member_id], 2, 'Yes' if member.completed else 'No')

            column = 3
            for part in badge.parts:
                worksheet.write(1, column, part.name, bold_format)
                part.badge = badge_map[part.id]
                if not part.badge.progress_loaded:
                    part.badge.load_progress(self._conn)
                    print('-> Loaded "%s"...' % (part.badge.name,))
                part_progress = {}

                part_count = 1
                groups = {}
                if part.group:
                    last_part = ''
                    part_count = 0
                    for badge_part in part.badge.parts:
                        this_part = badge_part.name.strip()
                        if last_part != this_part:
                            last_part = this_part
                            worksheet.write(2, column + part_count, last_part, bold_format)
                            groups[last_part] = []
                            part_count += 1
                        groups[last_part].append(badge_part.part_id)
                else:
                    groups['all'] = len(part.badge.parts)

                print('--> Calculating progress')
                for item in part.badge.progress:
                    if item.completed:
                        part_progress[item.member_id] = [1.0 for _ in range(part_count)]
                    else:
                        if part.group:
                            part_progress[item.member_id] = [0.0 for _ in range(part_count)]
                            pos = 0
                            for grp_list in groups.values():
                                count = 0
                                for id in grp_list:
                                    if id in item.parts:
                                        count += 1
                                actual_progress = count / len(grp_list)
                                part_progress[item.member_id][pos] = actual_progress
                                pos += 1
                        else:
                            total = len(item.parts)
                            part_progress[item.member_id] = [total / groups['all']]

                print('--> Exporting')
                for member in member_map.keys():
                    row = member_map[member]
                    item_progress = part_progress[member]
                    for loop in range(part_count):
                        worksheet.write(row, column + loop, item_progress[loop], progress_format)

                column += part_count
        
            last_column = xlsxwriter.utility.xl_col_to_name(column - 1)
            last_row = str(len(members) + (3 if badge.group else 2))
            range_to_format = ('D4' if badge.group else 'D3') + ':' + last_column + last_row
            worksheet.conditional_format(range_to_format, 
                {
                    'type': 'icon_set',
                    'icon_style': '5_arrows',
                    'icons': [
                        {'criteria': '>', 'type': 'number', 'value': 0.99},
                        {'criteria': '>', 'type': 'number', 'value': 0.75},
                        {'criteria': '>', 'type': 'number', 'value': 0.50},
                        {'criteria': '>', 'type': 'number', 'value': 0.25}
                    ]
                })


def ensureExtension(filename, extension):
    return filename if filename.lower().endswith(extension) else filename + extension

if __name__ == "__main__":
    mgr = ReportGenerator()
    mgr.run()
