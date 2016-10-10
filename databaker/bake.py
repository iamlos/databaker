#!/usr/bin/python

"""
Usage:
  bake.py [options] <recipe> <spreadsheet> [<params>...]

Options:
  --notiming            Suppress timing information.
  --preview             Preview selected cells in Excel.
  --nocsv               Don't produce CSV file.
  --debug               Debug Mode
  --nolookuperrors      Dont output 'NoLookuperror' to final CSV.
"""

import atexit
import codecs
import imp
import re
import sys
sys.stdout = codecs.getwriter('utf-8')(sys.stdout)

from timeit import default_timer as timer

from docopt import docopt
import xypath
import xypath.loader
import os.path
from technicalcsv import TechnicalCSV

import overrides        # warning: changes xypath and messytables
import warnings
import xlutils.copy
import xlwt
import richxlrd.richxlrd as richxlrd
from datetime import datetime
import string

# If there's a custom template, use it. Otherwise use the default.
try:
    import structure_csv_user as template
    from structure_csv_user import *
except ImportError:
    import structure_csv_default as template
    from structure_csv_default import *


__version__ = "1.0.7"
Opt = None
crash_msg = []


def showtime(msg='unspecified'):
    if not Opt or not Opt.timing:
        return
    global last
    t = timer()
    print "{}: {:.3f}s,  {:.3f}s total".format(msg, t - last, t - start)
    last = t

def onexit():
    return showtime('exit')

start = timer()
last = start




class Options(object):
    def __init__(self):
        options = docopt(__doc__, version='databaker {}'.format(__version__))
        self.xls_files = [options['<spreadsheet>']]
        self.recipe_file = options['<recipe>']
        self.timing = not options['--notiming']
        self.preview = options['--preview']
        self.preview_filename = "preview-{spreadsheet}-{recipe}-{params}.xls"
        self.csv_filename = "data-{spreadsheet}-{recipe}-{params}.csv"
        self.csv = not options['--nocsv']
        self.debug = options['--debug']
        self.no_lookup_error = not options['--nolookuperrors']
        self.params = options['<params>']



class Progress(object):
    # creates a progress bar
    def __init__(self, max_count, prefix=None, msg="\r{}{:3d}% - [{}{}]"):
        self.last_percent = None
        self.max_count = max_count
        self.msg = msg
        if prefix is not None:
            self.prefix = prefix + ' - '
        else:
            self.prefix = ''

    def update(self, count):
        percent = (((count+1) * 100) // self.max_count)
        if percent != self.last_percent:
            progress = percent / 5
            print self.msg.format(self.prefix, percent, '='*progress, " "*(20-progress)),
            sys.stdout.flush()
            self.last_percent = percent

def per_file(spreadsheet, recipe):
    def filenames():
        get_base = lambda filename: os.path.splitext(os.path.basename(filename))[0]
        xls_directory = os.path.dirname(spreadsheet)
        xls_base = get_base(spreadsheet)
        recipe_base = get_base(Opt.recipe_file)
        parsed_params = ','.join(Opt.params)

        csv_filename = Opt.csv_filename.format(spreadsheet=xls_base,
                                               recipe=recipe_base,
                                               params=parsed_params)

        csv_path = os.path.join(xls_directory, csv_filename)

        preview_filename = Opt.preview_filename.format(spreadsheet=xls_base,
                                                       recipe=recipe_base,
                                                       params=parsed_params)
        preview_path = os.path.join(xls_directory, preview_filename)
        return {'csv': csv_path, 'preview': preview_path}

    def make_preview():
        # call for each segment
        for i, header in tab.headers.items():
            if hasattr(header, 'bag') and not isinstance(header.bag, xypath.Table):
                for bag in header.bag:
                    writer.get_sheet(tab.index).write(bag.y, bag.x, bag.value,
                        xlwt.easyxf('pattern: pattern solid, fore-colour {}'.format(colourlist[i])))
                for ob in segment:
                    writer.get_sheet(tab.index).write(ob.y, ob.x, ob.value,
                        xlwt.easyxf('pattern: pattern solid, fore-colour {}'.format(colourlist[OBS])))


    tableset = xypath.loader.table_set(spreadsheet, extension='xls')
    showtime("file {!r} imported".format(spreadsheet))
    if Opt.preview:
        writer = xlutils.copy.copy(tableset.workbook)
    if Opt.csv:
        csv_file = filenames()['csv']
        csv = TechnicalCSV(csv_file, Opt.no_lookup_error)
        
    tabs = list(xypath.loader.get_sheets(tableset, recipe.per_file(tableset)))
    if not tabs:
        print "No matching tabs found."
        exit(1)
    bheaderoutput = False
    for tab_num, tab in enumerate(tabs):
        showtime("tab {!r} imported".format(tab.name))
        
        ## The callback into the recipe
        try:
            pertab = recipe.per_tab(tab)
            
            
            
        except Exception:
            crash_msg.append("tab: {!r} {!r}".format(tab_num, tab.name))
            raise
            
        # process the per_tab return value
        if isinstance(pertab, xypath.xypath.Bag):
            pertab = [pertab]
        #print("jjjk")
        #print(csv.generate_header_row(tab))

        try:
            for seg_id, segment in enumerate(pertab):
                if Opt.debug:
                    print "tab and segment available for interrogation"
                    import pdb; pdb.set_trace()

                if Opt.preview:
                    make_preview()

                if Opt.csv and len(segment) != 0:
                    obs_count = len(segment)
                    progress = Progress(obs_count, 'Tab {}'.format(tab_num + 1))
                    
                    csv.begin_observation_batch(tab)
                    if not bheaderoutput:
                        csv.csv_writer.writerow(csv.generate_header_row(tab))
                        bheaderoutput = True
                        
                    for ob_num, ob in enumerate(segment):  
                        assert ob.table == tab
                        try:
                            csv.handle_observation(ob)
                        except Exception:
                            crash_msg.append("ob: {!r}".format(ob))
                            raise
                        progress.update(ob_num)
                    print()
                    csv.finish_observation_batch()
                    
                # hacky observation wiping
                tab.headers = {}
                tab.max_header = 0
                tab.headernames = [None]
        except Exception:
            crash_msg.append("segment: {!r}".format(seg_id))
            crash_msg.append("tab: {!r} {!r}".format(tab_num, tab.name))
            raise


    if Opt.csv:
        csv.footer()
    if Opt.preview:
        writer.save(filenames()['preview'])

def create_colourlist():
    # Function to dynamically assign colours to dimensions for preview
    "https://github.com/python-excel/xlwt/blob/master/xlwt/Style.py#L309"
    colours = ["lavender", "violet", "gray25", "sea_green",
              "pale_blue", "blue", "gray25", "rose", "tan", "light_yellow", "light_green", "light_turquoise",
              "light_blue", "sky_blue", "plum", "gold", "lime", "coral", "periwinkle", "ice_blue", "aqua"]
    numbers = []
    for i in range(len(template.dimension_names)-1, \
                   -(len(colours) - len(template.dimension_names)), -1):
        numbers.append(-i)
    colourlist = dict(zip(numbers, colours))
    return colourlist
colourlist = create_colourlist()



def main():
    global Opt
    Opt = Options()
    atexit.register(onexit)
    recipe = imp.load_source("recipe", Opt.recipe_file)
    for fn in Opt.xls_files:
        try:
            per_file(fn, recipe)
        except Exception:
            crash_msg.append("fn: {!r}".format(fn))
            crash_msg.append("recipe: {!r}".format(Opt.recipe_file))
            print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
            print '\n'.join(crash_msg)
            print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
            raise

if __name__ == '__main__':
    main()
