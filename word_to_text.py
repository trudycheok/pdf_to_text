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


def get_text_d2p(filename):
    result = docx2python(filename)
    string_list = [i for i in flatten_list(result.document)]
    return remove_placeholders(string_list)


def get_text_wp():
    from win32com import client

    doc_folder = r"word_to_text"
    word_app = client.Dispatch("Word.Application")
    word_app.Visible = False
    wb = word_app.Documents.Open(f"{doc_folder}/sample.docx")

    text_list = []
    for w in range(1, wb.Words.Count + 1):
        text_list.append(str(wb.Words(w))
                         .replace("\x02", "")
                         .replace("\x07", "")
                         .replace("\r", "\n"))

    print("".join(text_list))

    table_list = []

    for t in range(1, wb.Tables.Count + 1):
        for c in range(1, wb.Tables(1).Columns.Count + 1):
            for r in range(1, wb.Tables(1).Rows.Count + 1):
                text = (wb.Tables(t).Cell(Row=r, Column=c).Range.Text
                        .replace("\r", "")
                        .replace("\x07", ""))
                table_list.append(text)

    print(table_list)

    footnotes_list = []
    for i in range(1, wb.Footnotes.Count + 1):
        footnotes_list.append(wb.Footnotes.Item(i).Range.Text)

    print(footnotes_list)


if __name__ == '__main__':
    # print(get_text_d2p("sample.docx"))
    print(get_text_wp())