import yagmail
import mimetypes
import io
import os
import sys

from tabulate import tabulate
from os.path import isfile

PY3 = sys.version_info[0] > 2


def check_for_inline_images(contents):
    attachments = []
    for i in range(len(contents)):
        item = contents[i]
        if isinstance(item, str) and isfile(item):
            if mimetypes.guess_type(item)[0].split("/")[0] == "image":
                contents[i] = yagmail.inline(item)
            else:
                attachments.append(item)
    for item in attachments:
        contents.remove(item)
    return contents, attachments


def prepare_contents(contents=None, attachments=None, prettify_html=True):
    """Prepares the contents of the mail. Identifies inline images inside of a textbody in contents (they have to be put in their own line). Identifies attachments inside of contents. Checks if attachments exist.

    Raises:
        TypeError: If path in attachments don't exist.

    Returns:
        list or str, list or str: The contents and attachments of the mail
    """
    if not isinstance(contents, (list, tuple)):
        if contents is not None:
            contents = [contents]
    if not isinstance(attachments, (list, tuple)):
        if attachments is not None:
            attachments = [attachments]
    if attachments is None:
        attachments = []
    if contents is None:
        contents = []
    subcontents_list = []
    indices = []
    # The following code block enables one to include images in a string textblock
    for item in contents:
        if not isinstance(item, (list, tuple, dict, set)):
            if isfile(item):
                attachments.append(item)
                contents.remove(item)
    for i in range(len(contents)):
        if not isinstance(contents[i], (list, tuple, dict, set)):
            # split the string at every new line to check if there is a path in one line
            subcontents = contents[i].split("\n")
            subcontents = [x.strip() for x in subcontents]
            if len(subcontents) > 1:
                subcontents_list.append(subcontents)
                indices.append(i)
    for index, subcontents in zip(indices, subcontents_list):
        contents.pop(index)
        contents[index:index] = subcontents
    contents, extra_attachments = check_for_inline_images(contents)
    attachments += extra_attachments
    if attachments is not None:
        for a in attachments:
            if isinstance(a, str):
                if not isinstance(a, io.IOBase) and not isfile(a):
                    raise TypeError(
                        f"{a} must be a valid filepath or file handle (instance of io.IOBase). {a} is of type {type(a)}"
                    )
    return contents, attachments


def get_message_html_and_str(contents, prettify_html=True):
    """Creates the plain text and hmtl body according to given contents. Note that contents first has to be prepared using def prepare_contents

    Args:
        contents (list): The result of prepare_contents
        prettify_html (bool, optional): If set to True, will use premailer to transform the html code. Defaults to True.

    Returns:
        str, str: The plain text and html body
    """
    htmlstr = ""
    plainstr = ""
    for item in contents:
        if type(item) == yagmail.inline:
            alias = os.path.basename(str(item))
            hashed_ref = str(abs(hash(alias)))
            htmlstr += '<img src="cid:{0}" title="{1}"/>'.format(hashed_ref, alias)
            plainstr += "-- img {0} should be here -- ".format(alias) + "\n"
        elif type(item) == list:
            htmlstr += list_to_html_table(item)
            plainstr += tabulate(item[1:], headers=item[0]) + "\n"
        else:
            try:
                htmlstr += "<div>{0}</div>".format(item)
                if PY3 and prettify_html:
                    import premailer

                    htmlstr = premailer.transform(htmlstr)
            except UnicodeEncodeError:
                htmlstr += "<div>{0}</div>".format(item)
            plainstr += item + "\n"
    return plainstr, htmlstr


def list_to_html_table(table, style=None):
    """Turns a list into a html table

    Args:
        table (list):
        style (str, optional): The html style of the table. Defaults to None.

    Returns:
        str: Html code for the table
    """
    htmlstr = ""
    rowstr = ""
    if style is None:
        htmlstr = "<style>table, th, td {border:1px solid black;}</style>"
    for cell in table[0]:
        cell = str(cell)
        rowstr += "<th>{}</th>".format(cell)
    htmlstr += "<tr>{}</tr>".format(rowstr)
    for row in table[1:]:
        rowstr = ""
        for cell in row:
            cell = str(cell)
            rowstr += "<td>{}</td>".format(cell)
        htmlstr += "<tr>{}</tr>".format(rowstr)
    htmlstr = "<table style = 'width:100%'>{}</table>".format(htmlstr)
    return htmlstr
