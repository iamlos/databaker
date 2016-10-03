import re
import xypath
import warnings
from utf8csv import UnicodeWriter

# If there's a custom template, use it. Otherwise use the default.
try:
    import structure_csv_user as template
    from structure_csv_user import *
except ImportError:
    import structure_csv_default as template
    from structure_csv_default import *

def rewrite_headers(row,dims):
    for i in range(0,len(row)):
        if i >= len(template.start.split(',')):
            which_cell_in_spread = (i - len(template.start.split(','))) % len(template.value_spread)
            which_dim = (i - len(template.start.split(','))) / len(template.value_spread)
            which_dim = int(which_dim)
            if value_spread[which_cell_in_spread] == 'value':
                row[i] = dims[which_dim]
    return row



def dim_name(dimension):
    if isinstance(dimension, int) and dimension <= 0:
        # the last dimension is dimension 0; but we index it as -1.
        return template.dimension_names[dimension-1]
    else:
        return dimension

def parse_ob(ob):
    if isinstance(ob.value, datetime):
        return (ob.value, '')
    if isinstance(ob.value, float):
        return (ob.value, '')
    if ob.properties['richtext']:
        string = richxlrd.RichCell(ob.properties.cell.sheet, ob.y, ob.x).fragments.not_script.value
    else:
        string = ob.value
    value, datamarker = re.match(r"([-+]?[0-9]+\.?[0-9]*)?(.*)", string).groups()
    if value is None:
        value = ''
    return value.strip(), datamarker.strip()


def datematch(date, silent=False):
    """match mmm yyyy, mmm-mmm yyyy, yyyy Qn, yyyy"""
    if not isinstance(date, basestring):
        if isinstance(date, float) and date>=1000 and date<=9999 and int(date)==date:
            return "Year"
        if not silent:
            warnings.warn("Couldn't identify date {!r}".format(date))
        return ''
    d = date.strip()
    if re.match('\d{4}$', d):
        return 'Year'
    if re.match('\d{4} [Qq]\d$', d):
        return 'Quarter'
    if re.match('[A-Za-z]{3}-[A-Za-z]{3} \d{4}$', d):
        return 'Quarter'
    if re.match('[A-Za-z]{3} \d{4}$', d):
        return 'Month'
    if not silent:
        warnings.warn("Couldn't identify date {!r}".format(date))
    return ''


class TechnicalCSV(object):
    def __init__(self, filename, no_lookup_error):
        self.no_lookup_error = no_lookup_error
        self.filehandle = open(filename, "wb")
        self.csv_writer = UnicodeWriter(self.filehandle)
        self.row_count = 0
        self.header_dimensions = None

    def write_header_if_needed(self, dimensions, ob):
        if self.header_dimensions is not None:
            # we've already written headers.
            return
        self.header_dimensions = dimensions
        header_row = template.start.split(',')

        # create new header row
        for i in range(dimensions):
            header_row.extend(template.repeat.format(num=i+1).split(','))

        # overwrite dimensions/subject/name as column header (if requested)
        if template.topic_headers_as_dims:
            dims = []
            for dimension in range(1, ob._cell.table.max_header+1):
                dims.append(ob._cell.table.headernames[dimension])
            header_row = rewrite_headers(header_row, dims)

        # Write to the file
        self.csv_writer.writerow(header_row)


    def footer(self):
        self.csv_writer.writerow(["*"*9, str(self.row_count)])
        self.filehandle.close()

    def output(self, row):
        def translator(s):
            if not isinstance(s, basestring):
                return unicode(s)
            # this is slow. We can't just use translate because some of the
            # strings are unicode. This adds 0.2 seconds to a 3.4 second run.
            return unicode(s.replace('\n',' ').replace('\r', ' '))
        self.csv_writer.writerow([translator(item) for item in row])
        self.row_count += 1

    def cell_for_dimension(self, obj, dimension):
        try:
            cell = obj.table.headers.get(dimension, lambda _: None)(obj)
        except xypath.xypath.NoLookupError:
            print "no lookup to dimension {} from cell {}".format(dim_name(dimension), repr(obj))
            if self.no_lookup_error:
                cell = "NoLookupError"            # if user wants - output 'NoLookUpError' to CSV
            else:
                cell = ''                         # Otherwise output a blanks
        return cell

    def value_for_dimension(self, obj, dimension):
        # implicit: obj
        cell = self.cell_for_dimension(obj, dimension)
        if cell is None:
            value = ''
        elif isinstance(cell, (basestring, float)):
            value = cell
        elif cell.properties['richtext']:
            value = richxlrd.RichCell(cell.properties.cell.sheet, cell.y, cell.x).fragments.not_script.value
        else:
            value = cell.value
        return value

    def get_dimensions_for_ob(self, ob):

        # TODO not really 'self'y
        """For a single observation cell, provide all the
           information for a single CSV row"""
        out = {}
        obj = ob._cell
        keys = ob.table.headers.keys()


        # Get fixed headers.
        values = {}
        values[OBS] = obj.value

        LAST_METADATA = 0 # since they're numbered -9 for obs, ... 0 for last one
        for dimension in range(OBS+1, LAST_METADATA + 1):
            values[dimension] = self.value_for_dimension(obj, dimension)

        # Mutate values
        # Special handling per dimension.
        # NOTE  - variables beginning SH_ ... are dependent on user choices from the template file

        if template.SH_Split_OBS:
            if not isinstance(values[OBS], float):  # NOTE xls specific!
                ob_value, dm_value = parse_ob(ob)
                values[OBS] = ob_value
                # the observation is not actually a number
                # store it as a datamarker and nuke the observation field
                if values[template.SH_Split_OBS] == '':
                    values[template.SH_Split_OBS] = dm_value
                elif dm_value:
                    logging.warn("datamarker lost: {} on {!r}".format(dm_value, ob))

        if template.SH_Create_ONS_time:
            if values[TIMEUNIT] == '' and values[TIME] != '':
                # we've not actually been given a timeunit, but we have a time
                # determine the timeunit from the time
                values[TIMEUNIT] = datematch(values[TIME])

        for dimension in range(OBS, LAST_METADATA + 1):
            yield values[dimension]
            if dimension in template.SH_Repeat:         # Calls special handling - repeats
                yield values[dimension]
            for i in range(0, template.SKIP_AFTER[dimension]):
                yield ''

        for dimension in range(1, obj.table.max_header+1):
            name = obj.table.headernames[dimension]
            value = self.value_for_dimension(obj, dimension)
            topic_headers = template.get_topic_headers(name, value)
            for col in topic_headers:
                yield col

    def handle_observation(self, ob):
        number_of_dimensions = ob.table.max_header
        self.write_header_if_needed(number_of_dimensions, ob)
        output_row = self.get_dimensions_for_ob(ob)
        self.output(output_row)
