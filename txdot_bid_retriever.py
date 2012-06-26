#!/usr/bin/env python

# A tool for downloading bid status from the Texas Department of Transportation
# and generating an Excel spreadsheet.

# Written by Charles Duffy <charles@dyfis.net> and submitted to the public
# domain.

import sys
import webapp2

try:
	from cStringIO import StringIO
except ImportError:
	from StringIO import StringIO
import xlwt
import lxml.html
import urlparse

from google.appengine.api import urlfetch

def html_from_url(url):
	"""Work around Google App Engine pecularities"""
	f = urlfetch.fetch(url)
	if f.status_code != 200:
		raise RuntimeException("Status code %r retrieving %r" % (f.status_code, url))
	return lxml.html.parse(StringIO(f.content))

MONTHS = {
	'January': 1,
	'February': 2,
	'March': 3,
	'April': 4,
	'May': 5,
	'June': 6,
	'July': 7,
	'August': 8,
	'September': 9,
	'October': 10,
	'November': 11,
	'December': 12,
}

def sheet_name_for_link(link_name):
	retval = link_name.replace(' Let ', ' ').replace('Projects for ', '- ')
	for month_name, month_num in MONTHS.iteritems():
		retval = retval.replace('%s, ' % month_name, '%s.' % month_num)
	return retval

def build_sheet():
	main_page_url = 'http://www.txdot.gov/business/bt.htm'
	main_page = html_from_url(main_page_url)

	links = []
	for link in main_page.xpath('//a[contains(@href, "/bidtab/")][contains(@href, ".htm")]'):
		links.append((link.text, link.attrib['href']))

	wb = xlwt.Workbook()
	for link_name, link_url_rel in links:
		sheet_name = sheet_name_for_link(link_name)
		print >>sys.stderr, '%s' % sheet_name
		ws = wb.add_sheet(sheet_name)
		link_url = urlparse.urljoin(main_page_url, link_url_rel)
		data_page = html_from_url(link_url)
		data_tables = data_page.xpath('//table[tr/td//text()="Let Date"]')
		if len(data_tables) == 0:
			print >>sys.stderr, 'No table with Let Date found'
			ws.write(1, 1, "ERROR: No table with Let Date found")
			continue
		if len(data_tables) > 1:
			print >>sys.stderr, 'Multiple tables with Let Date found'
			ws.write(1, 1, "ERROR: Multiple tables with Let Date found")
			continue
		data_table = data_tables[0]
		row_num = 0
		for tr_el in data_table:
			print >>sys.stderr, '%s,%s' % (sheet_name, row_num)
			col_num = 0
			if tr_el.tag != 'tr':
				print >>sys.stderr, 'Expected %r, found %r; skipping' % ('tr', tr_el.tag)
				continue
			for td_el in tr_el:
				if td_el.tag != 'td':
					print >>sys.stderr, 'Expected %r, found %r; skipping' % ('td', td_el.tag)
					continue
				print >>sys.stderr, '%s,%s,%s' % (sheet_name, row_num, col_num)
				content = ' '.join(td_el.itertext())
				print >>sys.stderr, '%s,%s,%s: %r' % (sheet_name, row_num, col_num, content)
				ws.write(row_num, col_num, ' '.join(td_el.itertext()))
				col_num+= 1
			row_num+= 1
		if row_num == 0:
			print >>sys.stderr, 'No rows found'
			ws.write(1, 1, 'ERROR: No rows found')
		else:
			print >>sys.stderr, 'Processed %r rows' % row_num

	outstream = StringIO()
	wb.save(outstream)
	return outstream.getvalue()

import webapp2

class MainPage(webapp2.RequestHandler):
	def get(self):
		sheet_text = build_sheet()
		self.response.headers['Content-Type'] = 'application/vnd.ms-excel'
		self.response.headers['Content-Disposition'] = 'attachment;filename="txdot_bids.xls"'
		self.response.out.write(sheet_text)
app = webapp2.WSGIApplication([('/', MainPage)], debug=True)
