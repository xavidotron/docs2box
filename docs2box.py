#!/usr/bin/python
# -*- coding: utf-8 -*-
from __future__ import print_function
import pickle
import os, shutil, io, re
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import pyPdf

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
# This needs Arial.ttf to be located next to this script. It's the default
# font used by Google Docs. Unfortunately, it's copyrighted and I can't just
# put it in the github repo. Helvetica is a built-in PDF font but isn't
# convenient in Google Docs.  Maybe I should standardize on Roboto or something.
pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

CHARS = {}
PACKET_STUFF = {}

SKIP_NOUNS = ('shard', 'runes', 'stone')

EXTRACT_RE = re.compile(ur'^(?:The )?["“]([^"]+)["”](?:, a)? (?:[Bb]luesheet|[Rr]itual)')

DUPLEX = {'char'}

OUTDIR = 'Fading Lights/'

def main():
    def download_as(item, path):
        request = service.files().export_media(fileId=item['id'],
                                               mimeType='application/pdf')
        fh = open(path, 'wb')
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            print("Download %d%% of %s (%s)." % (
                int(status.progress() * 100),
                item['name'], item['mimeType']))

    if os.path.exists(OUTDIR):
        shutil.rmtree(OUTDIR)
    if os.path.exists('Tmp/'):
        shutil.rmtree('Tmp/')
    os.mkdir(OUTDIR)
    os.mkdir(OUTDIR + 'Packets/')
    os.mkdir(OUTDIR + 'Printing/')
    os.mkdir('Tmp/')

    blank = 'Tmp/blank.pdf'
    c = canvas.Canvas(blank, pagesize=letter)
    c.setFont("Arial", 11)
    c.drawCentredString(letter[0]/2, 550, 'This page intentionally left blank.')
    c.showPage()
    c.save()
    
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('drive', 'v3', credentials=creds)

    # Call the Drive v3 API
    results = service.files().list(
        q=u"name='くコ:彡  L I M I N A L I T Y   S Q U I D  くコ:彡 / Fading Lights'",
        pageSize=10, fields="files(id, name, mimeType)").execute()
    items = results.get('files', [])

    if not items:
        print('No file found by that name.')
    elif len(items) > 1:
        print('Found multiple files:')
        for item in items:
            print(u'{0} ({1})'.format(item['name'], item['id']))
    else:
        folder = items[0]
        assert folder['mimeType'] == 'application/vnd.google-apps.folder', folder
        results = service.files().list(
            q="'%s' in parents" % folder['id'],
            pageSize=50, fields="files(id, name, mimeType)").execute()
        contents = results.get('files', [])
        for item in contents:
            if item['mimeType'] == 'application/vnd.google-apps.folder':
                os.mkdir('%s%s/' % (OUTDIR, item['name']))
                results = service.files().list(
                    q="'%s' in parents" % item['id'],
                    pageSize=50, fields="files(id, name, mimeType)").execute()
                subcontents = results.get('files', [])
                for j in subcontents:
                    assert j['mimeType'] == 'application/vnd.google-apps.document', j
                    path = '%s%s/%s.pdf' % (OUTDIR, item['name'], j['name'])
                    download_as(j, path)
                    if item['name'].startswith('Char'):
                        trequest = service.files().export_media(
                            fileId=j['id'],
                            mimeType='text/plain')
                        tfh = io.BytesIO()
                        tdownloader = MediaIoBaseDownload(tfh, trequest)
                        done = False
                        while done is False:
                            status, done = tdownloader.next_chunk()
                            print("Text %d%%." % int(status.progress() * 100))
                        stufflist = False
                        stuff = []
                        value = unicode(tfh.getvalue(), 'utf-8')
                        for l in value.split('\n'):
                            if '.' not in l and ('Sheets' in l or 'Items' in l):
                                stufflist = True
                            elif not l.startswith('* '):
                                stufflist = False
                            elif stufflist:
                                stuff.append(l[2:])
                        CHARS[path] = stuff

                        if ', ' in j['name']:
                            headerName = j['name'].split(', ', 1)[0]
                        else:
                            headerName = j['name']
                        c = canvas.Canvas("Tmp/%s.pdf" % j['name'],
                                          pagesize=letter)
                        c.setFont("Arial", 11)
                        c.drawRightString(letter[0]-inch,
                                          letter[1]-.65*inch, headerName)
                        c.showPage()
                        c.save()
                    else:
                        PACKET_STUFF[j['name']] = path
            elif item['name'].endswith(' (x)'):
                print('Skippping %s.' % item['name'])
            elif item['mimeType'] == 'application/vnd.google-apps.document':
                download_as(item, '%s%s.pdf' % (OUTDIR, item['name']))

    extracted = dict((p, 0) for p in PACKET_STUFF)
    skipped = dict((s, 0) for s in SKIP_NOUNS)
    colors = {'char': [], 'blue': [], 'green': [], 'white': []}
    for c in CHARS:
        writer = pyPdf.PdfFileWriter()
        reader = pyPdf.PdfFileReader(open(c))
        outpages = []
        for p in reader.pages:
            writer.addPage(p)
            outpages.append(p)
        colors['char'].append(outpages)
        name_dot_pdf = c.rsplit('/', 1)[-1]
        for s in CHARS[c]:
            m = EXTRACT_RE.search(s)
            if m:
                assert m.group(1) in PACKET_STUFF, (s, m.group(1))
                extracted[m.group(1)] += 1
                sheet = PACKET_STUFF[m.group(1)]
                outpages = []
                reader = pyPdf.PdfFileReader(open('Tmp/%s' % name_dot_pdf))
                namep = reader.pages[0]
                reader = pyPdf.PdfFileReader(open(sheet))
                for p in reader.pages:
                    p.mergePage(namep)
                    outpages.append(p)
                    writer.addPage(p)
                color = None
                for col in colors:
                    if sheet.lower().startswith(OUTDIR.lower() + col):
                        color = col
                        break
                else:
                    color = 'white'
                colors[color].append(outpages)
            else:
                for skip in SKIP_NOUNS:
                    if skip in s.lower():
                        skipped[skip] += 1
                        break
                else:
                    assert False, repr(s)
        writer.write(open('%sPackets/%s' % (OUTDIR, name_dot_pdf), 'w'))
    print('Skipped:', skipped)
    print('Extracted:', extracted)

    blankreader = pyPdf.PdfFileReader(open(blank))
    blankp = blankreader.pages[0]
    for col in colors:
        if len(colors[col]) > 0:
            pagecnt = 0
            writer = pyPdf.PdfFileWriter()
            for outpages in colors[col]:
                for p in outpages:
                    writer.addPage(p)
                    pagecnt += 1
                if col in DUPLEX and pagecnt % 2 == 1:
                    writer.addPage(blankp)
                    pagecnt += 1
            writer.write(open('%sPrinting/%s.pdf' % (OUTDIR, col), 'w'))

if __name__ == '__main__':
    main()
    
