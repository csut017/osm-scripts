''' Script for downloading the badge progress for a division in a term. '''

from osm import Connection, Manager

def main():
    ''' Main script. '''
    print('Connecting')
    conn = Connection('secret.json')
    conn.connect()

    print('Downloading sections and terms')
    mgr = Manager()
    mgr.load(conn)

    print('The following sections and terms are available:')
    for section in mgr.sections:
        print('   %s=>%s' % (section.section_id, str(section)))
        for term in section.terms:
            print('      %s=>%s' % (term.term_id, str(term)))

    print('Downloading badges')
    for section in mgr.sections:
        print(str(section))
        term = section.terms[-1]
        term.load_badges(conn)
        for badge in term.badges:
            print('   %s' % (badge.name))

        badge = term.find_badge('Gold Kea Award')
        print('Gold Kea Award')
        for part in badge.parts:
            print('   %s' % (part.name))
        badge.load_progress(conn)
        badge.export_progress('test.xlsx')


if __name__ == "__main__":
    main()
