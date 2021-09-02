""" Testing module for easemail """

import itertools
import easemail
from easemail import client


def get_combinations(ema):
    """ Creates permutations of possible inputs """
    tos = (
        None,
        (ema.user),
        [ema.user, ema.user],
        {ema.user: '"me" <{}>'.format(ema.user), ema.user + "1": '"me" <{}>'.format(ema.user)},
    )
    subjects = ("subj", ["subj"], ["subj", "subj1"])
    contents = (
        None,
        ["body"],
        ["body", "body1", "<h2><center>Text</center></h2>", u"<h1>\u2013</h1>"],
        [[["col1", "col2"], [1, 1], ["text", "cell"]]],
    )
    results = []
    for row in itertools.product(tos, subjects, contents):
        options = {y: z for y, z in zip(["to", "subject", "contents"], row)}
        options["display_only"] = True
        results.append(options)

    return results


def test_one():
    """ Tests several versions of allowed input for an email address """
    ema = easemail.client("test@yahoo.com", smtp_skip_login=True, soft_email_validation=False)
    mail_combinations = get_combinations(ema)
    for combination in mail_combinations:
        print(ema.send(**combination))


def test_two():
    """ Tests several versions of allowed input for outlook """
    ema = easemail.client("outlook", smtp_skip_login=True, soft_email_validation=False)
    mail_combinations = get_combinations(ema)
    for combination in mail_combinations:
        print(ema.send(print_only=True, **combination))
