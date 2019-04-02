import os
import sys

from datetime import date
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Cm, Pt
from osm import AwardScheme, Connection, Manager

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

        print 'Retrieving badge report...'
        report = self._term.load_badges_by_person(self._conn)

        print 'Generating report...'
        filename = ensureExtension(sys.argv[2]+'-Badge Report', '.docx')
        document = Document()
        self._generate_report(report, document)

        print 'Saving...'
        document.save(filename)

        print 'Done'

    def _set_term(self, args):
        term = args[0]
        self._set_section(args[1:])

        print 'Setting term...'
        if term == 'current':
            term = self._section.current_term()
            if term is None:
                print '-> Currently not in a term'
                return
            else:
                self._term = term
                print '-> Term set to %s' % (str(term), )
                return

        for term in self._section.terms:
            if term.name == term:
                self._term = term
                print '-> Term set to %s' % (str(term), )
                return
            
        print '-> Unknown term: %s' % (term, )

    def _set_section(self, args):
        print 'Setting section...'
        section = self._mgr.find_section(args[0])
        if section is None:
            print '-> Unknown section: %s' % (args[0], )
        else:
            self._section = section
            print '-> Section set to %s' % (str(section), )
    
    def _generate_report(self, report, document):
        headingStyle = document.styles.add_style('TableHeading', WD_STYLE_TYPE.PARAGRAPH)
        headingStyle.font.bold=True

        section = document.sections[0]
        new_width, new_height = section.page_height, section.page_width
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = new_width
        section.page_height = new_height
        section.left_margin = Cm(1)
        section.right_margin = Cm(1)

        section.header.paragraphs[0].text = 'Badge Report'
        now = date.today()
        section.footer.paragraphs[0].text = 'Generated ' + now.strftime('%d %B %Y')
        table = document.add_table(rows = 1, cols = 2, style='Table Grid')
        table.columns[0].width = Cm(5)
        table.columns[1].width = Cm(21)
        cells = table.rows[0].cells
        cells[0].text = 'Person'
        cells[1].text = 'Badges'
        cells[0].width = Cm(5)
        cells[1].width = Cm(21)
        clearFormatting(cells[0].paragraphs[0], headingStyle)
        clearFormatting(cells[1].paragraphs[0], headingStyle)
        for person in report:
            cells = table.add_row().cells
            name = '%s %s' % (person.first_name, person.last_name)
            print '...adding row for %s...' % (name, )
            cells[0].text = name
            clearFormatting(cells[0].paragraphs[0])
            para = cells[1].paragraphs[0]
            clearFormatting(para)
            cells[0].width = Cm(5)
            cells[1].width = Cm(21)
            for badge in person.badges:
                badge_path = os.path.join('badge_images', os.path.basename(badge.picture))
                if not os.path.exists(badge_path):
                    print '...retrieving badge image for %s...' % (badge.name,)
                    self._conn.download_binary(badge.picture, badge_path)
                if badge.completed:
                    para.add_run().add_picture(badge_path, width = Cm(2))
                    para.add_run(' ')


def ensureExtension(filename, extension):
    return filename if filename.lower().endswith(extension) else filename + extension

def clearFormatting(paragraph, style=None):
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    if style is not None:
        paragraph.style = style

if __name__ == "__main__":
    mgr = ReportGenerator()
    mgr.run()
