from dataclasses import dataclass
from dateutil.relativedelta import relativedelta
import datetime
import math
import re

from misc import utils
from venues.errors import HashError, NoValidSessionsException

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
    
    def __hash__(self):
        # Hashing is done only with zone and street
        # number so that if some data is not formatted
        # in the same way, that's okay.
        try:
            return hash((
                self.zone,
                re.findall(r'[0-9]+', self.street)[0]
            ))
        except IndexError as e:
            raise HashError(f"The address '{self.street}' contains no number to use for hashing.")


    def __eq__(self, other: 'VenueRecord') -> bool:
        return self.__hash__() == other.__hash__()

    @property
    def average_rsvps(self) -> int:
        """Average number of RSVPs for all jobs with this venue,
        rounded up to the nearest whole number.
        """
        if len(self.job_records) == 0:
            return 0.0
        
        total_rsvps = sum(job.rvsps for job in self.job_records)
        return  math.ceil(total_rsvps / len(self.job_records))

    @property
    def average_ror(self) -> float:
        """Total number of RSVPs and RMIs across all
        jobs divided by total quantity across all jobs,
        and multiplied by 100 (to express as a percent).
        """
        total_rsvps_rmis = 0
        total_quantity = 0

        for job in self.job_records:
            total_rsvps_rmis += job.rvsps + job.rmi
            total_quantity += job.quantity
        
        if total_quantity == 0:
            return 0
        
        return 100 * total_rsvps_rmis / total_quantity

    @property
    def latest_job(self) -> 'JobRecord':
        latest_job: JobRecord = None
        
        for job in self.job_records:
            if (latest_job is None
                or latest_job.end_date < job.end_date):
                latest_job = job
        
        return latest_job

    @staticmethod
    def from_entry(entry: dict[str, str]) -> 'VenueRecord':
        """Create a new venue object from an entry dictionary.
        """
        new_venue = VenueRecord(
            VenueRecord.strip_field(entry['MKT']),
            int(entry['LOC#']),
            VenueRecord.strip_field(entry['Zone']),
            VenueRecord.strip_field(entry['Restaurant']),
            VenueRecord.strip_field(entry['St Address']),
            VenueRecord.strip_field(entry['City']),
            VenueRecord.strip_field(entry['ST']),
            int(entry['ZIP']))

        new_venue.add_job_record(entry)
        
        return new_venue
    
    @staticmethod
    def strip_field(field: str | None) -> str:
        """Return a stripped `str` if field is not `None`,
        else returns an empty `str`.
        """
        return field.strip() if field is not None else ''
    
    # IMPORTANT the header order MUST match
    # the order of data in to_entry()'s returned tuple.
    # there is no mechanism checking if they match.
    def to_entry(self) -> tuple[str]:
        """Returns a spreadsheet-ready representation of this venue.
        """
        return (
            self.latest_job.id,
            self.latest_job.user,
            self.market,
            self.loc_num,
            self.latest_job.week,
            self.zone,
            self.restaurant,
            self.street,
            self.city,
            self.state,
            self.zip,
            "Menu",  # Rod wants everything to say Menu
            self.latest_job.month,
            self.latest_job.year,
            self.latest_job.num_sessions,
            self.latest_job.session_type,
            self.latest_job.quantity,
            self.latest_job.rvsps,
            self.latest_job.rmi,
            self.latest_job.ror,
            self.average_rsvps,
            self.average_ror
        )
    
    def within(self, weeks: int, ref_date: datetime) -> bool:
        """Returns `True` if at least one session in a job took
        place within a certain number of weeks from `ref_date`.

        E.g., `SomeVenue.within(16, datetime.now())` would return `True`
        if `SomeVenue` had a job within 16 weeks of today.
        """
        threshold_date = ref_date - relativedelta(weeks=weeks)
        for job in self.job_records:
            for session in job.sessions:
                if session.datetime > threshold_date:
                    return True
        
        return False
    
    def around_time_last_year(self, start_date: datetime, end_date: datetime, prox_weeks: int) -> bool:
        """Returns `True` if at least one session in a job took place
        during the time last year between `start_date` and `end_date` or within `prox_weeks` weeks thereof.
        """
        start_threshold = start_date - relativedelta(years=1) - relativedelta(weeks=prox_weeks)
        end_threshold = end_date - relativedelta(years=1) + relativedelta(weeks=prox_weeks)

        for job in self.job_records:
            for session in job.sessions:
                if (session.datetime >= start_threshold and
                    session.datetime <= end_threshold):
                    return True
        
        return False

    def add_job_record(self, entry: dict[str, str]) -> None:
        """Create a job record for `entry` and add it to this venue's job records
        if the job entry is valid.
        """
        new_job = JobRecord.from_entry(entry)
        self.job_records.add(new_job)
        

@dataclass(frozen=True)
class JobRecord:
    """A record of a job for a particular venue. A job record only contains
    information about the job itself and does not contain information about the
    venue.
    """
    id: int
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

    def __hash__(self):
        return hash((
            self.id,
            self.user,
            self.week,
            self.mail_piece,
            self.month,
            self.year,
            self.num_sessions,
            self.quantity,
            self.rvsps,
            self.rmi
        ))

    def __eq__(self, other: 'JobRecord') -> bool:
        return self.__hash__() == other.__hash__()
    
    @property
    def month_date(self) -> datetime:
        return utils.parse_month_year(self.month, self.year)
    
    @property
    def end_date(self) -> datetime:
        return self.latest_session.datetime

    @property
    def latest_session(self) -> 'SessionRecord':
        latest_session: SessionRecord = None
        for session in self.sessions:
            if (latest_session is None
                or latest_session.datetime < session.datetime):
                latest_session = session
        
        return latest_session
    
    @property
    def session_type(self) -> str:
        lunches = 0
        dinners = 0
        for session in self.sessions:
            if session.meal_type == 'Lunch':
                lunches += 1
            else:
                dinners += 1
        
        return f'{lunches} Lunch {dinners} Dinner'
    
    @property
    def ror (self) -> float:
        """The Rate of Return. This is the total number of RSVPs
        and RMIs (request for more information) divided by job quantity and
        multiplied by 100 (to express as a percent).
        """
        if self.quantity == 0:
            return 0
        
        return 100 * (self.rvsps + self.rmi) / self.quantity
    
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

@dataclass(frozen=True)
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
                except (ValueError, TypeError):
                    continue

                new_session = SessionRecord(
                    meal_type,
                    entry[day_of_week_key],
                    date_and_time)
                    
                sessions.add(new_session)
        
        if len(sessions) == 0:
            raise NoValidSessionsException('No valid sessions found in entry.')
        
        return sessions