from datetime import datetime, timedelta


def convertNonStringKeys(dict_):
    return {str(key): val for key, val in dict_.items()}


def convertBackToDateKeys(dict_):
    return {datetime.strptime(key, "%Y-%m-%d"): val for key, val in dict_.items()}


def get_dates_range(first, length):
    return [first + timedelta(day) for day in range(length)]