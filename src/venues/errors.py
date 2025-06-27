class NoValidSessionsException(Exception):
    """Entry contains no valid sessions.
    """
    pass

class HashError(Exception):
     """Object data not hashable.
     """
     pass