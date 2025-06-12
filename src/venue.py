import datetime
from functools import cached_property

class Venue:

    venue_list: list['Venue'] = []

    def __init__(self, entry, add_to_list=False):
        assert entry is not None, 'Cannot create a venue with no entries'
        self.entries: list[dict[str, str]] = [entry]
        if (add_to_list):
            Venue.venue_list.append(self)

    def add_entry(self, entry: dict[str, str]) -> 'Venue':
        """
        Adds a new entry (instance of event) to a venue's list of entries.
        """
        self.entries.append(entry)
        return self
    
    @staticmethod
    def not_in_list(venue: 'Venue') -> bool:
        return venue.address not in [other_venues.address for other_venues in Venue.venue_list]
    
    def within_four_months(self, ref_date: datetime.datetime) -> bool:
        """
        Returns True if any entry for this venue has a date within the last four months of ref_date.
        """
        def parse_date(date_str):
            for fmt in ("%m/%d/%y", "%m-%d-%y"):
                try:
                    return datetime.datetime.strptime(date_str, fmt)
                except (ValueError, TypeError):
                    continue
            return None

        cutoff = ref_date - datetime.timedelta(days=120)
        date_fields = [
            'Lunch 1 Date', 'Lunch 2 Date', 'Lunch 3 Date', 'Lunch 4 Date',
            'Dinner 1 Date', 'Dinner 2 Date', 'Dinner 3 Date'
        ]
        for entry in self.entries:
            for field in date_fields:
                date_val = parse_date(entry.get(field, ''))
                if date_val and cutoff <= date_val <= ref_date:
                    return True
        return False

    @cached_property
    def dict_repr(self) -> dict[str, str]:
        """Dictionary representation of this venue. This function is used when
        writing venues to the final csv file.
        """
        return {
            'MKT': self.attrib('MKT'),
            'Zone': self.attrib('Zone'),
            'Restaurant': self.attrib('Restaurant'),
            'St Address': self.attrib('St Address'),
            'City': self.attrib('City'),
            'ST': self.attrib('ST'),
            'ZIP': self.attrib('ZIP')
        }

    @cached_property
    def address(self) -> tuple:
        """
        Tuple representing the address of the venue in the form: (name, street, city, state, zip),
        """
        return (self.attrib('Restaurant'), self.attrib('Zone'), self.attrib('St Address'), self.attrib('City'), self.attrib('ST'), self.attrib('ZIP'))

    @cached_property
    def average_rsvps(self) -> float:
        sum = 0
        for entry in self.entries:
            sum += int(entry.get('RSVPs'))
        return sum / len(self.entries)

    def attrib(self, key: str) -> str:
        """
        Returns an attribute of this venue.
        """
        return self.entries[0].get(key)