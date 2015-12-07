#!/usr/bin/python

import os
import xlrd
import json
from jsonpath_rw import jsonpath, parse
import datetime
import urllib2

import httplib

# A utility to delete weights from journal based on excel spreadsheet.
# The spreadsheet contains UUID's, dates
# The script gets the weight ID from the identity endpoint and then calls the journal endpoint with UUID, date and weight id to delete the weight.
# "path" is the path to the spreadsheet where the entries are.

def process( path ):

    if os.path.isfile( output_file ):
        os.remove( output_file )
    file = open( output_file, 'a' )

    book=xlrd.open_workbook( path )
    sheet = book.sheet_by_name( sheet_name )

    number_rows = sheet.nrows
    #conn = httplib.HTTPConnection('localhost:9900')

    for idx in range( start_row, number_rows ):
        if sheet.cell( idx, 2).value != '':
            oracle_weight_date=convert_date( sheet.cell( idx, 3).value, book)
            uuid = sheet.cell( idx, 2).value
            json = get( uuid, oracle_weight_date )
            list_weight_date_id = parse_json( uuid, oracle_weight_date, json )

            for idx,tup in enumerate( list_weight_date_id ):
                write_to_file( file, uuid, list_weight_date_id[idx][0], list_weight_date_id[idx][1], list_weight_date_id[idx][2] )
                delete( uuid,list_weight_date_id[idx] )

    file.close
    #conn.close


def write_to_file( file, uuid, weight, date, id ):
    file.write( uuid + ",\t" + str(weight) + "\t" + str(date) + "\t\n")


def get( uuid, date ):
    url = "http://localhost:9900/journal/%s/weight/daily/%s?count=20&units=kgs" % (uuid,date)
    try:
        url_contents = urllib2.urlopen( url )
        result = json.load( url_contents )
        print "Got JSON(%s,%s): %s" % (uuid,date,result)

    except:
        print "No entry for (uuid,date)=(%s,%s)" % (uuid, date)
        result = {}

    return result


def delete( uuid, weight_date_id ):

    url = "/journal/%s/weight/%s/%s" % (uuid, weight_date_id[1].replace("-",""), weight_date_id[2] )
    print "DELETING: %s" % (url)
    conn = httplib.HTTPConnection('localhost:9900')
    conn.request('DELETE', url)
    resp = conn.getresponse()
    conn.close()
    print ("(%s): DELETED: %s" % (resp.status, url ) )


# "24-Nov-15" to "20151123"
def convert_date( date, book ):

    asdf = datetime.datetime(*xlrd.xldate_as_tuple(date, book.datemode))

    return asdf.strftime( "%Y%m%d")


def parse_json( uuid, date, json_string ):
    print "Parsing Json: %s" % json_string
    weights = [match.value for match in parse( '[*].weight' ).find( json_string )]
    dates = [match.value for match in parse( '[*].date' ).find( json_string )]
    ids = [match.value for match in parse( '[*].id' ).find( json_string )]

    print "weightparse=%s" % weights
    length = len( weights )
    result = list()
    for idx,weight in enumerate( weights ):
        if date == str( dates[idx] ).replace("-",""):
            result.append((weight, dates[idx], ids[idx]))

    if len(result) >1 :
        print( "WARN:%s HAS %s weights for %s" % (uuid, len(result), date))
    return result


#----------------------------------------------------------------------
if __name__ == "__main__":
    #path = "/Users/david.lu/Downloads/weight_diff.xlsx"
    path = "/Users/david.lu/Desktop/asdf.xlsx"
    sheet_name="Delete CORE"
    output_file="core_delete.txt"
    start_col=0
    start_row=3-1
    process(path)

