'''
This script generates an Excel spreadsheet of the attendence at the division.
'''

import os
import sys

from datetime import date
import xlsxwriter
from osm import Connection, Manager

import matplotlib.pyplot as plt


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
        print 'Connecting to OSM...'
        self._connect()
        self._initialise()

        if len(sys.argv) < 3:
            print 'ERROR: term and section have not been set! '
            return

        self._set_term(sys.argv[1:3])
        if self._term is None:
            return

        print 'Retrieving term programme...'
        report = self._term.load_programme(self._conn, True)

        print 'Generating report...'
        filename = ensureExtension(sys.argv[2]+'-Attendence', '.xlsx')
        workbook = xlsxwriter.Workbook(filename)
        self._generate_report(report, workbook)

        print 'Saving to %s...' % (filename, )
        workbook.close()

        print 'Done'

    def _set_term(self, args):
        term_name = args[0]
        self._set_section(args[1:])

        print 'Setting term...'
        if term_name == 'current':
            term = self._section.current_term()
            if term is None:
                print '-> Currently not in a term'
                return
            else:
                self._term = term
                print '-> Term set to %s' % (str(term), )
                return
        
        for term in self._section.terms:
            if term.name == term_name:
                self._term = term
                print '-> Term set to %s' % (str(term), )
                return
            
        print '-> Unknown term: %s' % (term_name, )

    def _set_section(self, args):
        print 'Setting section...'
        section = self._mgr.find_section(args[0])
        if section is None:
            print '-> Unknown section: %s' % (args[0], )
        else:
            self._section = section
            print '-> Section set to %s' % (str(section), )
    
    def _generate_report(self, report, workbook):
        bold = workbook.add_format({'bold': True, 'font_size': 12})
        for meeting in report:
            print '-> Processing meeting ' + meeting.name + ' on ' + meeting.date.strftime('%d-%m-%Y')
            ws_name = meeting.date.strftime('%d-%m-%Y') + ' ' + meeting.name
            if len(ws_name) > 30:
                ws_name = ws_name[0:27] + '...'
            worksheet = workbook.add_worksheet(ws_name)
            worksheet.write('A1', meeting.name, bold)
            worksheet.write('A2', 'First Name', bold)
            worksheet.write('B2', 'Family Name', bold)
            worksheet.write('C2', 'Patrol', bold)
            row = 2
            for member in meeting.members:
                worksheet.write(row, 0, member.first_name)
                worksheet.write(row, 1, member.last_name)
                worksheet.write(row, 2, member.patrol)
                row += 1


def ensureExtension(filename, extension):
    return filename if filename.lower().endswith(extension) else filename + extension

if __name__ == "__main__":
    mgr = ReportGenerator()
    mgr.run()
