from optparse import OptionParser
from MarketSurvey import MarketSurvey

usage = "usage: %prog [options] arg"
parser = OptionParser(usage)
parser.add_option("-f", "--file", dest="filename",
                  help="read data from market survey FILENAME", metavar="FILE")
parser.add_option("-v", "--verbose",
                  action="store_true", dest="verbose")

(options, args) = parser.parse_args()

s = MarketSurvey()

if options.filename:
    if options.verbose:
        print "Reading %s..." % options.filename
    s.init(filename=options.filename)

if s.wb:
    if options.verbose:
        print "%s opened successfully" % options.filename
    if s.ws:
        if options.verbose:
            print "current worksheet is %s" % s.ws.title
        # Check to see if there is text in the District Name
        s.get_district()
        if s.district:
            if options.verbose:
                print "found district: %s" % s.district
        s.get_province()
        if s.province:
            if options.verbose:
                print "found province: %s" % s.province
        s.get_date()
        if s.date:
            if options.verbose:
                print "found date: %s" % s.date
        s.get_collector()
        if s.collector:
            if options.verbose:
                print "found collector: %s" % s.collector
        # Roll up all the crops:
        for row_number in range(14, 80):
            if s.is_crop(row_number):
                s.parse_row(row_number)
                if options.verbose:
                    if s.crops[-1]['avg_prices'][0]:
                        print "found crop: %s with average prices %f, %f" % (
                            s.crops[-1]['name'],
                            s.crops[-1]['avg_prices'][0],
                            s.crops[-1]['avg_prices'][1]
                        )
