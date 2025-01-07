# utils/date_formatter.py
from datetime import datetime

class TurkishDateFormatter:
    MONTH_MAP = {
        1: 'Oca',
        2: 'Şub',
        3: 'Mar',
        4: 'Nis',
        5: 'May',
        6: 'Haz',
        7: 'Tem',
        8: 'Ağu',
        9: 'Eyl',
        10: 'Eki',
        11: 'Kas',
        12: 'Ara'
    }

    @classmethod
    def format_date(cls, date: datetime = None) -> str:
        """Format date in Turkish MMM.YY format"""
        if date is None:
            date = datetime.now()
        
        month_tr = cls.MONTH_MAP[date.month]
        year = str(date.year)[2:]  # Get last 2 digits of year
        
        return f"{month_tr}.{year}"