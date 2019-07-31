''' Objects for working with OSM data. '''

from datetime import datetime
import csv
import json
import requests
import xlsxwriter


class Connection(object):
    ''' Connection to OSM. '''

    def __init__(self, settings_path):
        with open(settings_path) as f:
            settings = json.load(f)
            self._server = settings['server']
            self._token = settings['token']
            self._api_id = settings['apiID']
            self._username = settings['userName']
            self._password = settings['password']
        self._user_id = None
        self._secret = None

    def connect(self):
        ''' Connects to the server. '''
        data = {
            'token': self._token,
            'apiid': self._api_id,
            'email': self._username,
            'password': self._password
        }
        req = requests.post(
            self._server + '/users.php?action=authorise', data=data)
        resp = req.json()
        try:
            self._user_id = resp['userid']
            self._secret = resp['secret']
        except KeyError:
            raise Error('Unable to connect: ' + resp['error'])

    def download(self, url):
        ''' Downloads some data from the server. '''
        data = {
            'token': self._token,
            'apiid': self._api_id,
            'userid': self._user_id,
            'secret': self._secret
        }
        req = requests.post(self._server + url, data=data)
        req.raise_for_status()
        return req.json()

    def upload(self, url, data):
        ''' Downloads some data from the server. '''
        data['token'] = self._token
        data['apiid'] = self._api_id
        data['userid'] = self._user_id
        data['secret'] = self._secret
        req = requests.post(self._server + url, data=data)
        req.raise_for_status()
        if req.text != '':
            return req.json()
        return {}

    def download_binary(self, url, filename):
        ''' Downloads a binary file from the server'''
        url = url if url.startswith('/') else u'/' + url
        req = requests.get(self._server + str(url))
        with open(filename, 'wb') as f:
            f.write(req.content)


class Error(Exception):
    ''' Connection errors. '''

    def __init__(self, message):
        super(Error, self).__init__()
        self.message = message

    def __str__(self):
        return self.message


class Manager(object):
    ''' OSM data. '''

    def __init__(self):
        self.sections = []

    def load(self, conn):
        ''' Loads the data for a manager. '''
        data = conn.download('/api.php?action=getUserRoles')
        sections = {}
        for rec in data:
            section = Section(rec)
            sections[section.section_id] = section
            self.sections.append(section)

        data = conn.download('/api.php?action=getTerms')
        for key, value in data.items():
            section = sections[key]
            for rec in value:
                term = Term(rec, section)
                section.terms.append(term)

    def find_section(self, name=None, section_type=None):
        ''' Finds a section. '''
        if section_type is None:
            sections = [
                section for section in self.sections if section.name == name]
        else:
            sections = [
                section for section in self.sections if section.type == section_type]

        try:
            return sections[0]
        except IndexError:
            return None


class Section(object):
    ''' A scouting section. '''

    def __init__(self, source):
        self.name = source["sectionname"]
        self.type = source["section"]
        self.group = source["groupname"]
        self.section_id = source["sectionid"]
        self.terms = []
        self.badges = []

    def __str__(self):
        return '%s: %s [%s]' % (self.group, self.name, self.type)

    def current_term(self):
        now = datetime.now().date()
        for term in self.terms:
            if term.start_date <= now and term.end_date >= now:
                return term
        return None


class Term(object):
    ''' A term within a section. '''

    def __init__(self, source, section):
        self.name = source["name"]
        self.start_date = datetime.strptime(
            source["startdate"], '%Y-%m-%d').date()
        self.end_date = datetime.strptime(source["enddate"], '%Y-%m-%d').date()
        self.term_id = source["termid"]
        self.section = section
        self.badges = []
        self.badges_loaded = False
        self.programme = []
        self.programme_loaded = 0
        self.members = []
        self.members_loaded = False

    def __str__(self):
        start_date = self.start_date.strftime('%Y-%m-%d')
        end_date = self.end_date.strftime('%Y-%m-%d')
        return '%s (%s to %s)' % (self.name, start_date, end_date)

    def load_badges(self, conn):
        '''Retrieves the badges for the term. '''
        self.badges = []
        self.__load_badges(conn, 1)
        self.__load_badges(conn, 2)
        self.badges_loaded = True

    def load_badges_by_person(self, conn):
        '''Retrieves the badges for the members for the term. '''
        data = conn.download('/ext/badges/badgesbyperson/?action=loadBadgesByMember&sectionid=%s&term_id=%s' %
                             (self.section.section_id, self.term_id))
        badge_report = []
        for rec in data['data']:
            member = Member(rec)
            badge_report.append(member)
        return badge_report

    def __load_badges(self, conn, badge_type):
        '''Retrieves the badges for the term. '''
        number = len(self.badges) + 1
        data = conn.download(
            '/ext/badges/records/?action=getBadgeStructureByType' +
            '&a=1&section=%s&type_id=%s&term_id=%s&section_id=%s' %
            (self.section.type, badge_type, self.term_id, self.section.section_id))
        details = data['details']
        structure = data['structure']
        for _, value in details.items():
            badge_id = value['badge_identifier']
            try:
                parts = structure[badge_id]
            except KeyError:
                parts = [{}, {}]
            badge = Badge(number, value, parts, self)
            self.badges.append(badge)
            number += 1

    def find_badge(self, name):
        ''' Finds a badge by its name. '''
        badges = [badge for badge in self.badges if badge.name == name]
        try:
            return badges[0]
        except IndexError:
            return None

    def load_programme(self, conn, include_attendance=False):
        ''' Loads the programme for the term. '''
        data = conn.download('/ext/programme/?action=getProgrammeSummary&sectionid=%s&termid=%s' %
                             (self.section.section_id, self.term_id))
        self.programme = []
        for rec in data['items']:
            meeting = Meeting(self, rec)
            self.programme.append(meeting)
        self.programme_loaded = 1

        if include_attendance:
            meetings = list([(meeting.date.strftime('%Y-%m-%d'), meeting)
                             for meeting in self.programme])
            data = conn.download('/ext/members/attendance/?action=get&sectionid=%s&termid=%s' %
                                 (self.section.section_id, self.term_id))
            for rec in data['items']:
                member = Member(rec)
                for meeting in meetings:
                    try:
                        if rec[meeting[0]] == 'Yes':
                            meeting[1].members.append(member)
                    except KeyError:
                        pass
            self.programme_loaded = 2

    def load_members(self, conn):
        ''' Loads the current members in the term. '''
        data = {
            'section_id': self.section.section_id,
            'term_id': self.term_id
        }
        data = conn.upload(
            '/ext/members/contact/grid/?action=getMembers', data)
        self.members = []
        for _, rec in data['data'].items():
            member = Member(rec)
            self.members.append(member)
        self.members_loaded = True

    def import_programme(self, filename, conn):
        ''' Imports a programme from a CSV file.
            This will update any existing programme. '''
        meetings = {}
        for meeting in self.programme:
            key = meeting.date.strftime('%y%m%d')
            meetings[key] = meeting

        with open(filename) as csvfile:
            csv_reader = csv.reader(csvfile)
            row = 0
            for data in csv_reader:
                row += 1
                if row > 1:
                    date = datetime.strptime(data[0], '%d-%b-%y').date()
                    key = date.strftime('%y%m%d')
                    try:
                        meeting = meetings[key]
                    except KeyError:
                        meeting = Meeting(self)
                        meeting.date = date
                    meeting.name = data[3]
                    meeting.parent_notes = data[4]
                    meeting.pre_notes = data[5]
                    meeting.leader = data[6]
                    if data[1] != '':
                        meeting.start_time = datetime.strptime(
                            data[1], '%H:%M')
                    if data[2] != '':
                        meeting.end_time = datetime.strptime(data[2], '%H:%M')
                    meeting.save(conn)


class Badge(object):
    ''' Defines a badge. '''

    def __init__(self, number, details, structure, term):
        self.number = number
        self.term = term
        self.section = term.section
        self.badge_id = details['badge_identifier']
        self.__id = details['badge_id']
        self.__version = details['badge_version']
        self.name = details['name']
        self.type = details['group_name']
        self.picture = details['picture']
        self.progress = []
        self.progress_loaded = False
        try:
            parts = structure[1]['rows']
        except (KeyError, IndexError):
            parts = []
        self.parts = list([BadgePart(part) for part in parts])

    def __str__(self):
        badge_type = self.type
        if badge_type is None or badge_type == '':
            badge_type = 'Unknown'
        return '%d: %s [%s]' % (self. number,
                                self.name,
                                badge_type)

    def load_progress(self, conn):
        ''' Loads the progress of the section for this badge. '''
        self.progress = []
        data = conn.download(
            '/ext/badges/records/?action=getBadgeRecords' +
            '&term_id=%s&section=%s&badge_id=%s&section_id=%s&badge_version=%s' %
            (self.term.term_id, self.section.type, self.__id,
             self.section.section_id, self.__version))
        for person in data["items"]:
            self.progress.append(BadgeProgress(person, self))
        self.progress_loaded = True

    def export_progress(self, filename=None, workbook=None):
        ''' Exports the badge progress to an Excel file. '''
        close_workbook = False
        if workbook is None:
            close_workbook = True
            workbook = xlsxwriter.Workbook(filename)

        ws_name = self.name
        if len(ws_name) > 30:
            ws_name = ws_name[0:27] + '...'
        worksheet = workbook.add_worksheet(ws_name)
        bold = workbook.add_format({'bold': True, 'font_size': 16})
        if len(self.parts) > 1:
            worksheet.merge_range(0, 0, 0, len(self.parts), self.name, bold)
        bold = workbook.add_format({'bold': True})
        worksheet.write('A2', 'Name', bold)
        col = 1
        for part in self.parts:
            worksheet.write(1, col, part.name, bold)
            col += 1
        row = 1
        for person in self.progress:
            row += 1
            col = 0
            worksheet.write(row, 0, person.firstname + ' ' + person.lastname)
            for part in self.parts:
                col += 1
                try:
                    worksheet.write(row, col, person.parts[part.part_id])
                except KeyError:
                    pass
        if close_workbook:
            workbook.close()


class BadgePart(object):
    ''' Defines a part of achieving the badge. '''

    def __init__(self, source):
        self.part_id = source['field']
        self.name = source['name']
        try:
            self.description = source['tooltip']
        except KeyError:
            self.description = ''

    def __str__(self):
        return self.name


class BadgeProgress(object):
    ''' Defines the progress towards a badge. '''

    def __init__(self, source, badge):
        self.badge = badge
        self.firstname = source['firstname']
        self.lastname = source['lastname']
        self.completed = source['completed'] == '1'
        self.parts = {}
        for part in badge.parts:
            part_id = part.part_id
            try:
                self.parts[part_id] = source[part_id]
            except KeyError:
                pass

    def __str__(self):
        completed = ' [Complete]' if self.completed else ''
        return '%s, %s%s' % (self.lastname, self.firstname, completed)


class Meeting(object):
    ''' Defines a meeting in a programme. '''

    def __init__(self, term, source=None):
        self.term = term
        self.members = []
        if source is None:
            self.name = None
            self.pre_notes = None
            self.post_notes = None
            self.parent_notes = None
            self.leader = None
            self.date = None
            self.start_time = None
            self.end_time = None
            self.meeting_id = None
        else:
            self.name = source['title']
            self.pre_notes = source['prenotes']
            self.post_notes = source['postnotes']
            self.parent_notes = source['notesforparents']
            self.leader = source['leaders']
            self.date = datetime.strptime(
                source["meetingdate"], '%Y-%m-%d').date()
            self.start_time = datetime.strptime(
                source["starttime"], '%H:%M:%S').time()
            self.end_time = datetime.strptime(
                source["endtime"], '%H:%M:%S').time()
            self.meeting_id = source["eveningid"]
            self.__save_state()

    def __save_state(self):
        self.__name = self.name
        self.__pre_notes = self.pre_notes
        self.__post_notes = self.post_notes
        self.__parent_notes = self.parent_notes
        self.__date = self.date
        self.__start_time = self.start_time
        self.__end_time = self.end_time
        self.__leader = self.leader

    def __str__(self):
        date = self.date.strftime('%Y-%m-%d')
        start_time = self.start_time.strftime('%I:%M%p')
        end_time = self.end_time.strftime('%I:%M%p')
        times = ''
        if start_time != '12:00AM' and end_time != '12:00AM':
            times = ' (' + start_time + ' to ' + end_time + ')'
        return '%s: %s%s' % (date, self.name, times)

    def save(self, conn):
        ''' Saves this meeting to OSM. '''
        if self.meeting_id is None:
            self.__add(conn)
        else:
            self.__update(conn)

    def __add(self, conn):
        ''' Adds a new meeting. '''
        data = {
            'sectionid': self.term.section.section_id,
            'start': self.date.strftime('%Y-%m-%d'),
            'title': self.name,
            'prenotes': self.pre_notes,
            'postnotes': self.post_notes,
            'leaders': self.leader,
            'notesforparents': self.parent_notes,
            'meetingdate': self.date.strftime('%Y-%m-%d'),
            'starttime': self.start_time.strftime('%H:%M'),
            'endtime': self.end_time.strftime('%H:%M')
        }
        data = conn.upload('/ext/programme/?action=addMeeting', data)
        self.meeting_id = data["lastmeetingadded"]
        self.__update(conn)

    def __update(self, conn):
        ''' Updates an existing meeting. '''
        parts = {}
        if self.name != self.__name:
            parts['title'] = self.name
        if self.pre_notes != self.__pre_notes:
            parts['prenotes'] = self.pre_notes
        if self.post_notes != self.__post_notes:
            parts['postnotes'] = self.post_notes
        if self.leader != self.__leader:
            parts['leaders'] = self.leader
        if self.parent_notes != self.__parent_notes:
            parts['notesforparents'] = self.parent_notes
        if self.date != self.__date:
            parts['meetingdate'] = self.date.strftime('%Y-%m-%d')
        if self.start_time != self.__start_time:
            parts['starttime'] = self.start_time.strftime('%H:%M')
        if self.end_time != self.__end_time:
            parts['endtime'] = self.end_time.strftime('%H:%M')
        self.__update_data(conn, parts)
        self.__save_state()

    def __update_data(self, conn, parts):
        data = {
            'sectionid': self.term.section.section_id,
            'termid': self.term.term_id,
            'eveningid': self.meeting_id,
            'parts': json.dumps(parts)
        }
        conn.upload('/ext/programme/?action=editEveningParts', data)

    def delete(self, conn):
        ''' Deletes this meeting from OSM. '''
        conn.download('/ext/programme/?action=deleteMeeting&eveningid=%s&sectionid=%s' %
                      (self.meeting_id, self.term.section.section_id))
        self.meeting_id = None


class Member(object):
    ''' Defines a member. '''

    def __init__(self, source):
        try:
            self.member_id = source['member_id']
        except KeyError:
            try:
                self.member_id = source['scout_id']
            except KeyError:
                self.member_id = source['scoutid']

        try:
            self.first_name = source['first_name']
        except KeyError:
            self.first_name = source['firstname']

        try:
            self.last_name = source['last_name']
        except KeyError:
            self.last_name = source['lastname']

        self.is_active = source['active']
        try:
            self.date_of_birth = source['date_of_birth']
        except KeyError:
            self.date_of_birth = source['dob']

        self.patrol = source['patrol']
        self.role = source['patrol_role_level_label']
        self.badges = []
        try:
            self.badges = [BadgeLink(badge) for badge in source['badges']]
        except KeyError:
            pass

    def __str__(self):
        return '%s, %s [%s]' % (self.last_name,
                                self.first_name,
                                (self.patrol + ' ' + self.role).strip())


class BadgeLink(object):
    def __init__(self, source):
        self.completed = source['completed'] == '1'
        self.awarded = source['awarded'] == '1'
        self.picture = source['picture']
        self.name = source['badge']


class AwardScheme(object):

    def __init__(self, settings_path):
        self.badges = []
        with open(settings_path) as f:
            data = json.load(f)
            for badge in data['badges']:
                self.badges.append(AwardSchemeBadge(badge))


class AwardSchemeBadge(object):

    def __init__(self, data):
        self.name = data['name']
        self.badge = None
        self.parts = []
        for part in data['parts']:
            self.parts.append(AwardSchemePart(part))


class AwardSchemePart(object):

    def __init__(self, data):
        self.name = data['name']
        self.id = data['id']
