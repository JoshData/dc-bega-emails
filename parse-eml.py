import datetime
import hashlib
import os
import sys
import subprocess

import html
import email.policy
from email import message_from_binary_file
from email.utils import parsedate_to_datetime

def save_file (fn, cont):
    """Saves cont to a file fn"""
    if os.path.exists(fn): return # we already extracted it
    with open(fn, "wb") as f:
        f.write(cont)
def construct_name (source_file_name, attachment_file_name):
    """Constructs a file name out of messages ID and packed file name"""
    name, ext = os.path.splitext(attachment_file_name)
    h = hashlib.sha1()
    h.update(source_file_name.encode("ascii"))
    h.update(attachment_file_name.encode("ascii"))
    return os.path.join("attachments", h.hexdigest() + ext)
def disqo (s):
    """Removes double or single quotations."""
    s = s.strip()
    if s.startswith("'") and s.endswith("'"): return s[1:-1]
    if s.startswith('"') and s.endswith('"'): return s[1:-1]
    return s
def disgra (s):
    """Removes < and > from HTML-like tag or e-mail address or e-mail ID."""
    s = s.strip()
    if s.startswith("<") and s.endswith(">"): return s[1:-1]
    return s
def pullout (m, source_file):
    """Extracts content from an e-mail message.
    This works for multipart and nested multipart messages too.
    m   -- email.Message() or mailbox.Message()
    key -- Initial message ID (some string)
    Returns tuple(Text, Html, Files, Parts)
    Text  -- All text from all parts.
    Html  -- All HTMLs from all parts
    Files -- Dictionary mapping extracted file to message ID it belongs to.
    Parts -- Number of parts in original message.
    """
    Html = ""
    Text = ""
    Files = {}
    Parts = 0
    if not m.is_multipart():
        if m.get_filename(): # It's an attachment
            fn = m.get_filename()
            cfn = construct_name(source_file, fn)
            Files[fn] = (cfn, None)
            content = m.get_content()
            if isinstance(content, str): content = content.encode("utf8")
            save_file(cfn, content)
            return Text, Html, Files, 1

        # Not an attachment!
        # See where this belongs. Text, Html or some other data:
        cp = m.get_content_type()
        if cp=="text/plain": Text += m.get_content()
        elif cp=="text/html": Html += m.get_content()
        else:
            # Something else!
            # Extract a message ID and a file name if there is one:
            # This is some packed file and name is contained in content-type header
            # instead of content-disposition header explicitly
            cp = m.get("content-type")
            try: id = disgra(m.get("content-id"))
            except: id = None
            # Find file name:
            o = cp.find("name=")
            if o==-1: return Text, Html, Files, 1
            ox = cp.find(";", o)
            if ox==-1: ox = None
            o += 5; fn = cp[o:ox]
            fn = disqo(fn)
            cfn = construct_name(source_file, fn)
            Files[fn] = (cfn, id)
            save_file(cfn, m.get_content())
        return Text, Html, Files, 1
    # This IS a multipart message.
    # So, we iterate over it and call pullout() recursively for each part.
    y = 0
    while 1:
        # If we cannot get the payload, it means we hit the end:
        try:
            pl = m.get_payload(y)
        except: break
        # pl is a new Message object which goes back to pullout
        t, h, f, p = pullout(pl, source_file)
        Text += t; Html += h; Files.update(f); Parts += p
        y += 1
    return Text, Html, Files, Parts

def extract (msgfile, source_file_name):
    """Extracts all data from e-mail, including From, To, etc., and returns it as a dictionary.
    msgfile -- A file-like readable object
    source_file_name     -- Some ID string for that particular Message. Can be a file name or anything.
    Returns dict()
    Keys: from, to, subject, date, text, html, parts[, files]
    Key files will be present only when message contained binary files.
    For more see __doc__ for pullout() and caption() functions.
    """

    if source_file_name.lower().endswith(".pdf"):
        From = None
        To = None
        Subject = source_file_name
        Date = None
        Text = subprocess.check_output(["pdftotext", source_file_name, "-"]).decode("utf8")
        Html = None
        Files = {}
    else:
        m = message_from_binary_file(msgfile, policy=email.policy.default)
        From, To, Subject, Date = caption(m, source_file_name)
        Text, Html, Files, Parts = pullout(m, source_file_name)
        Text = Text.strip(); Html = Html.strip()

    print("<article>")
    print("<h1>" + html.escape(Subject) + "</h1>")
    print("<div class='headers'>")
    if From: print("<p>From: " + html.escape(From) + "</p>")
    if To: print("<p>To: " + html.escape(To) + "</p>")
    if Date: print("<p>Date: <time datetime=\"" + str(Date) + "\">" + Date.strftime("%x %X") + "</time></p>")
    print("</div>")

    if Html:
        print(Html)
    else:
        print("<div style='white-space:pre-line;'>" + html.escape(Text) + "</div>")

    for attachment_name, (fn, id) in Files.items():
        print("<p><a href=\"" + fn + "\">" + attachment_name + "</a>")

    print("</article>")

def caption (origin, fn):
    """Extracts: To, From, Subject and Date from email.Message() or mailbox.Message()
    origin -- Message() object
    Returns tuple(From, To, Subject, Date)
    If message doesn't contain one/more of them, the empty strings will be returned.
    """
    Date = parsedate_to_datetime(origin["date"].strip().replace(" -0000", " +0000"))
    From = ""
    if "from" in origin: From = origin["from"].strip()
    To = ""
    if "to" in origin: To = origin["to"].strip()
    Subject = ""
    if "subject" in origin: Subject = origin["subject"].strip()
    return From, To, Subject, Date

def sort_by_date(fn):
    with open(fn, "rb") as f:
        if fn.lower().endswith(".pdf"):
            return datetime.datetime.now(datetime.timezone.utc) # TODO
        else:
            print(fn, file=sys.stderr)
            m = message_from_binary_file(f, policy=email.policy.default)
            From, To, Subject, Date = caption(m, fn)
            return Date

# Main.

# Get emails and put in order.
emails = sys.argv[1:]
emails.sort(key = sort_by_date)

# Dump in chronological order.
print("""<!DOCTYPE html>
<html lang="en-US">
  <head>
    <meta charset="utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
    <meta name="viewport" content="width=1024, user-scalable=no" />
    <style>
        article {
            margin: 1em 0;
            border: 1px solid #569;
            padding: 1em;
        }
        article .headers {
            color: #444;
            margin: 1em 0;
        }
        article .headers p {
            margin: 0;
        }
    </style>
  </head>
<body>""")

for fn in emails:
  with open(fn, "rb") as f:
    extract(f, f.name)

print("""</body>
</html>""")