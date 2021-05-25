from docx2python import docx2python
from collections.abc import Iterable
import itertools
import re


def flatten_list(lol):
    for l in lol:
        # String is an iterable, so need second condition
        if isinstance(l, Iterable) and not isinstance(l, (str, bytes)):
            # this mean for i in generator, yield i
            yield from flatten_list(l)
        else:
            yield l


def remove_placeholders(string_list):
    curr_list = [re.sub(r"----[\w]+\/?[\w]+\.?[\w]+----", "", string.strip()) for string in string_list]
    return [re.sub(r"footnote[0-9]+\)", "", string.strip()) for string in curr_list]


def get_text(filename):
    result = docx2python(filename)
    string_list = [i for i in flatten_list(result.document)]
    return remove_placeholders(string_list)


if __name__ == '__main__':
    print(get_text("sample.docx"))