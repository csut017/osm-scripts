''' Script for downloading and working with a term programme. '''

from osm import Connection, Manager


class ProgrammeManager(object):

    def __init__(self):
        self._conn = None
        self._mgr = None
        self._exit = False
        self._connect()
        self._initialise()
        self._section = None
        self._term = None

        self._commands = {
            'q': self._exit_manager,
            'quit': self._exit_manager,
            'exit': self._exit_manager,
            'sections': self._list_sections,
            'section': self._set_section,
            'terms': self._list_terms,
            'term': self._set_term,
        }


    def _connect(self):
        self._conn = Connection('secret.json')
        self._conn.connect()


    def _initialise(self):
        self._mgr = Manager()
        self._mgr.load(self._conn)


    def run(self):
        print 'Welcome to OSM programme manager'
        while not self._exit:
            cmd_text = raw_input('>').strip()
            cmd_args = self._split_cmd(cmd_text)
            if len(cmd_args) == 0:
                continue

            cmd_name = cmd_args[0]
            cmd_args = cmd_args[1:]
            try:
                cmd = self._commands[cmd_name]
            except KeyError:
                print 'Unknown command: %s' % (cmd_name, )
                continue

            try:
                cmd(cmd_args)
            except Exception as ex:
                print 'Unexpected error: %s' % (str(ex), )


    def _split_cmd(self, cmd_text):
        out = []
        word = ''
        in_string = False
        for c in cmd_text:
            if c == "'":
                in_string = not in_string
            elif c == ' ':
                if in_string:
                    word += c
                else:
                    if word != '':
                        out.append(word)
                    word = ''
            else:
                word += c

        if word != '':
            out.append(word)

        return out


    def _exit_manager(self, args):
        self._exit = True

    
    def _list_sections(self, args):
        for section in self._mgr.sections:
            print str(section)


    def _set_section(self, args):
        if len(args) < 1:
            print 'Current section is %s' % (str(self._section), )
            return

        section = self._mgr.find_section(args[0])
        if section is None:
            print 'Unknown section: %s' % (args[0], )
        else:
            self._section = section
            print 'Section set to %s' % (str(section), )

    
    def _list_terms(self, args):
        if self._section is None:
            print 'Section must be set first'
            return

        for term in self._section.terms:
            print str(term)


    def _set_term(self, args):
        if self._section is None:
            print 'Section must be set first'
            return

        if len(args) < 1:
            print 'Current term is %s' % (str(self._term), )
            return

        for term in self._section.terms:
            if term.name == args[0]:
                self._term = term
                print 'Term set to %s' % (str(term), )
                return
            
        print 'Unknown term: %s' % (args[0], )


def programme_print(mgr, conn):
    section = mgr.find_section(name='Cubs')
    print('Downloading programme for ' + str(section))
    term = section.terms[-1]
    term.load_programme(conn)
    for meeting in term.programme:
        print('   %s' % (str(meeting)))


if __name__ == "__main__":
    mgr = ProgrammeManager()
    mgr.run()
