from dataclasses import dataclass
import datetime

import misc.utils as utils

class VenueRecord:
    """A unique venue and its associated job records.
    """
    def __init__(self, market: str, loc_num: int, zone: str, restaurant: str, street: str, city: str, state: str, zip: int):
        self.market = market
        self.loc_num = loc_num
        self.zone = zone
        self.restaurant = restaurant
        self.street = street
        self.city = city
        self.state = state
        self.zip = zip
        self.job_records: set['JobRecord'] = set()
    
    def __eq__(self, other: 'VenueRecord') -> bool:
        return (self.market == other.market and
               self.loc_num == other.loc_num and
               self.zone == other.zone and
               self.restaurant == other.restaurant and
               self.street == other.street and 
               self.city == other.city and
               self.state == other.state and
               self.zip == other.zip)
    
    def __iter__(self) -> tuple['JobRecord']:
        return (job for job in self.job_records)

    @property
    def average_rsvps(self) -> float:
        """Average number of RSVPs for all jobs with this venue.
        """
        if len(self.job_records) == 0:
            return 0.0
        
        total_rsvps = sum(job.rvsps for job in self.job_records)
        return total_rsvps / len(self.job_records)

    @staticmethod
    def from_entry(entry: dict[str, str]) -> 'VenueRecord':
        """Create a new venue object from an entry dictionary.
        """
        new_venue = VenueRecord(
            entry['MKT'],
            int(entry['LOC#']),
            entry['Zone'],
            entry['Restaurant'],
            entry['St Address'],
            entry['City'],
            entry['ST'],
            int(entry['ZIP']))

        new_venue.add_job_record(entry)
        
        return new_venue

    def add_job_record(self, entry: dict[str, str]) -> None:
        """Create a job record for `entry` and add it to this venue's job records
        if the job entry is valid.
        """
        new_job = JobRecord.from_entry(entry)
        self.job_records.add(new_job)
        
    # TODO new implementation of method
    def within_four_months(self, ref_date: datetime.datetime) -> bool:
        pass
        

@dataclass
class JobRecord:
    """A record of a job for a particular venue. A job record only contains
    information about the job itself and does not contain information about the
    venue.
    """
    job_num: int
    user: str
    week: int
    mail_piece: str
    month: str
    year: int
    num_sessions: int
    sessions: set['SessionRecord']
    quantity: int
    rvsps: int
    rmi: int

    def __eq__(self, other: 'JobRecord') -> bool:
        return self.job_num == other.job_num
    
    def __iter__(self) -> tuple['SessionRecord']:
        return (session for session in self.sessions)
    
    @staticmethod
    def from_entry(entry: dict[str, str]) -> 'JobRecord':
        """Create a job record from an entry dictionary.
        Returns `None` if unable to create an the job.
        """
        new_job = JobRecord(
            int(entry['Job#']),
            entry['User'],
            int(entry['Week']),
            entry['Mail Piece'],
            entry['Month'],
            int(entry['Year']),
            int(entry['# Sessions']),
            SessionRecord.from_entry(entry),
            int(entry['Qty']),
            int(entry['RSVPs']),
            int(entry['RMI']))
            
        return new_job

@dataclass
class SessionRecord:
    """A record of a session for a particlar job. A session record only contains
    information about the session itself and does not contain information about the
    job."""
    meal_type: str
    day_of_week: str
    datetime: datetime.datetime

    def __eq__(self, other: 'SessionRecord') -> bool:
        return (self.meal_type == other.meal_type and
                self.day_of_week == other.day_of_week and
                self.datetime == other.datetime)
    
    @staticmethod
    def from_entry(entry: dict[str, str]) -> set['SessionRecord']:
        """Create a set of session records from an entry dictionary.
        """
        sessions: set['SessionRecord'] = set()

        for day in (1, 2, 3):
            for meal_type in ('Lunch', 'Dinner'):
                day_of_week_key  = f'{meal_type} Day {day}'
                date_key = f'{meal_type} {day} Date'
                time_key = f'{meal_type} {day} Time'

                # There will not be a session for every meal type and day,
                # so we just ignore if the datestring returns a ValueError
                # because the datestring is empty.
                try:
                    date_and_time = datetime.datetime.combine(entry[date_key], entry[time_key])
                except ValueError:
                    continue

                new_session = SessionRecord(
                    meal_type,
                    entry[day_of_week_key],
                    date_and_time)
                    
                sessions.add(new_session)
        
        if len(sessions) == 0:
            raise NoValidSessionsException('No valid sessions found in entry.')
        
        return sessions
    
class NoValidSessionsException(Exception):
    """Raised when no valid sessions are found in an entry.
    """
    pass