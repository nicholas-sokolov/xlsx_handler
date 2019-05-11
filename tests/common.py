import functools


def cases(case_list):
    def decorator(f):
        @functools.wraps(f)
        def wrapper(*args):
            for c in case_list:
                new_args = args + (c if isinstance(c, tuple) else (c,))
                f(*new_args)

        return wrapper

    return decorator
