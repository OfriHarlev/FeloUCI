#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
#    felo_rating.py - Core library of the Felo program
#
#    Copyright © 2006 Torsten Bronger <bronger@physik.rwth-aachen.de>
#
#    This file is part of the Felo program.
#
#    Felo is free software; you can redistribute it and/or modify it under
#    the terms of the MIT licence:
#
#    Permission is hereby granted, free of charge, to any person obtaining a
#    copy of this software and associated documentation files (the "Software"),
#    to deal in the Software without restriction, including without limitation
#    the rights to use, copy, modify, merge, publish, distribute, sublicense,
#    and/or sell copies of the Software, and to permit persons to whom the
#    Software is furnished to do so, subject to the following conditions:
#
#    The above copyright notice and this permission notice shall be included in
#    all copies or substantial portions of the Software.
#
#    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
#    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
#    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
#    THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
#    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
#    FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
#    DEALINGS IN THE SOFTWARE.
#

"""Core library for the Felo program, with the Felo numbers calculating code.

:author: Torsten Bronger
:copyright: 2006, Torsten Bronger
:license: MIT license
:contact: Torsten Bronger <bronger@physik.rwth-aachen.de>

:var datapath: the path to the directory where to expect the data files, in
  particular the auf*.dat files.  Normally, this is the same directory where
  this module was read from.
:var distribution_version: the version number of the current distribution of
  this module
:var separator: column separator in a Felo file
:var apparent_expectation_values: This list of lists is the solution for the
  winning-hit problem.  The winner of a bout is rated too highly because he
  ends the bout with his point.  This array contains the actually "measured"
  result values if the ideal expected result value is known (e.g. by using
  Elo's formula).

  The list contains 101 lists, for result values 0, 0.01, 0.02, ... 1.  (This
  happens to be the first element of each list, too.)  Every list contains 15
  values, for bouts fenced to 1, 2, ... 15 points.


:type datapath: string
:type distribution_version: string
:type separator: string
:type apparent_expectation_values: list
"""
__docformat__ = "restructuredtext en"

__all__ = ["Bout", "Fencer", "parse_felo_file", "write_felo_file", "calculate_felo_ratings",
           "expectation_value", "prognosticate_bout", "write_back_fencers",
           "write_back_fencers_to_file",
           "Error", "LineError", "BootstrappingError"]

__version__ = "$Revision: 213 $"
# $HeadURL: https://svn.sourceforge.net/svnroot/felo/src/felo_rating.py $
distribution_version = "1.0"

import codecs, re, os.path, datetime, time, shutil, glob, tempfile
# This strange construction is necessary because on Windows, the file may be
# put into a ZIP file (by py2exe), so we have to delete the last *two* parts of
# the path.
datapath = os.path.abspath(__file__)
while not os.path.isdir(datapath):
    datapath = os.path.dirname(datapath)
from subprocess import call, Popen, PIPE
import gettext
if os.name == 'nt':
    t = gettext.translation('felo', datapath+"/po", fallback=True)
else:
    t = gettext.translation('felo', fallback=True)
_ = t.ugettext


separator = "\s*\t+\s*"

def clean_up_line(line):
    """Removes leading and trailing whitespace and comments from the line.

    :Parameters:
      - `line`: the line to be cleaned-up

    :type line: string

    :Return:
      - the cleaned-up string

    :rtype: string
    """
    hash_position = line.find("#")
    if hash_position != -1:
        line = line[:hash_position]
    return line.strip()
    
def parse_items(input_file, linenumber=0):
    """Parses key--value pairs in the input_file.  It stops parsing if a line
    starts with a special character.  Normally, this is a line made of
    e.g. equal signs that acts as a separator between the sections in the Felo
    file.

    :Parameters:
      - `input_file`: an *open* file object where the items are read from
      - `linenumber`: number of already read lines in the input_file

    :type input_file: file
    :type linenumber: int

    :Return:
      - A dictionary with the key--value pairs, and the new number of already read
        lines.
      - The new current number of lines aready read in the file.

    :rtype: dict, int

    :Exceptions:
      - `LineError`: if a line has not the "name <TAB> value" format
    """
    items = {}
    line_pattern = re.compile("\s*(?P<name>[^\t]+?)"+separator+"(?P<value>.+?)\s*\\Z")
    for line in input_file:
        linenumber += 1
        line = clean_up_line(line)
        if not line: continue
        if line[0] in u"-=._:;,+*'~\"`´/\\%$!": break
        match = line_pattern.match(line)
        if not match:
            raise LineError(_('Line must follow the pattern "name <TAB> value".'), input_file.name, linenumber)
        name, value = match.groups()
        try:
            items[name] = value
            items[name] = float(value)
            items[name] = int(value)
        except ValueError:
            pass
    return items, linenumber

class Bout(object):
    """Class for single bouts.

    It is not very active in the sense that is has many methods.  Actually it's
    just a container for the attributes of a bout.

    :cvar date_pattern: regular expression pattern for parsing date strings.

    :type date_pattern: SRE_Pattern
    """
    date_pattern = re.compile("\s*(?P<year>\\d{4})-(?P<month>\\d{1,2})-(?P<day>\\d{1,2})"
                              "(?:\\.(?P<index>\\d+))")
    def __init__(self, year, month, day, index=0, first_fencer="", second_fencer="",
                 points_first=0, points_second=0, fenced_to=10):
        """Class constructor.

        :Parameters:
          - `year`: year of the bout
          - `month`: month of the bout
          - `day`: day of the bout
          - `index`: a number for bringing the bouts of a day in an order
          - `first_fencer`: name of the first fencer
          - `second_fencer`: name of the second fencer
          - `points_first`: points won by the first fencer
          - `points_second`: points won by the second fencer
          - `fenced_to`: winning points of the bout.  May be e.g. 5, 10, or 15.
            "0" means that the bout was part of a relay team competition.

        :type year: int
        :type month: int
        :type day: int
        :type index: int
        :type first_fencer: string
        :type second_fencer: string
        :type points_first: int
        :type points_second: int
        :type fenced_to: int
        """
        self.date = datetime.date(year, month, day)
        self.index = index
        self.first_fencer = first_fencer
        self.second_fencer = second_fencer
        self.points_first = int(points_first)
        self.points_second = int(points_second)
        self.fenced_to = int(fenced_to)
    def __cmp__(self, other):
        """If bouts are sorted with "sort()", they are sorted by their date,
        with early bouts first.
        """
        if self.date == other.date:
            return (self.date - other.date).total_seconds()
        else:
            return self.index - other.index
    def __get_date_string(self):
        result = self.date.isoformat()
        if self.index != 0:
            result += "." + str(self.index)
        return result
    def __set_date_string(self, date_string):
        match = Bout.date_pattern(date_string)
        if not match:
            raise FeloFormatError(_("Invalid date string."))
        year, month, day, index = match.groups("0")
        try:
            self.date = datetime.date(int(year), int(month), int(day))
            self.index = int(index)
        except (ValueError):
            raise FeloFormatError(_("Invalid date string."))
    date_string = property(__get_date_string, __set_date_string,
                           doc="""The date of the bout in its string representation YYYY-MM-TT.II,
                                  where ".II" is the optional index.
                                  @type: string""")

apparent_expectation_values = \
    [[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
     [0.01, 0.006766, 0.00751687, 0.00801342, 0.00834527, 0.00858217, 0.00875974, 0.00889779,
      0.0090082, 0.0090985, 0.00917373, 0.00923737, 0.00929191, 0.00933917, 0.00938052],
     [0.02, 0.013728, 0.0150747, 0.0160544, 0.0167146, 0.0171859, 0.017539, 0.0178135,
      0.0180328, 0.0182122, 0.0183616, 0.018488, 0.0185963, 0.0186901, 0.0187721],
     [0.03, 0.020882, 0.0226839, 0.0241246, 0.0251083, 0.0258114, 0.026338, 0.0267471,
      0.027074, 0.0273412, 0.0275637, 0.0277519, 0.0279131, 0.0280527, 0.0281748],
     [0.04, 0.028224, 0.030354, 0.0322269, 0.033527, 0.0344588, 0.0351568, 0.0356988,
      0.0361318, 0.0364856, 0.0367801, 0.0370291, 0.0372424, 0.0374271, 0.0375886],
     [0.05, 0.03575, 0.0380941, 0.0403646, 0.0419717, 0.0431285, 0.0439956, 0.0446687,
      0.0452063, 0.0456454, 0.0460109, 0.0463198, 0.0465843, 0.0468134, 0.0470136],
     [0.06, 0.043456, 0.0459125, 0.0485417, 0.0504437, 0.051821, 0.0528545, 0.0536569,
      0.0542975, 0.0548207, 0.055256, 0.0556239, 0.0559388, 0.0562115, 0.0564499],
     [0.07, 0.051338, 0.0538167, 0.056763, 0.0589451, 0.0605369, 0.0617339, 0.0626636,
      0.0634057, 0.0640116, 0.0645157, 0.0649415, 0.065306, 0.0656215, 0.0658973],
     [0.08, 0.059392, 0.0618139, 0.0650333, 0.067478, 0.0692772, 0.0706342, 0.0716889,
      0.0725309, 0.0732182, 0.0737899, 0.0742727, 0.0746859, 0.0750436, 0.0753561],
     [0.09, 0.067614, 0.0699104, 0.0733581, 0.0760453, 0.0780432, 0.0795558, 0.0807331,
      0.0816732, 0.0824406, 0.0830787, 0.0836175, 0.0840786, 0.0844776, 0.0848262],
     [0.1, 0.076, 0.078112, 0.0817429, 0.0846504, 0.0868365, 0.0884997, 0.0897966,
      0.0908328, 0.0916788, 0.0923821, 0.092976, 0.0934841, 0.0939236, 0.0943077],
     [0.11, 0.084546, 0.086424, 0.0901933, 0.0932969, 0.0956593, 0.0974669, 0.0988799,
      0.10001, 0.100933, 0.1017, 0.102348, 0.102902, 0.103382, 0.103801],
     [0.12, 0.093248, 0.094851, 0.0987153, 0.101989, 0.104514, 0.106459, 0.107984,
      0.109205, 0.110204, 0.111034, 0.111734, 0.112334, 0.112852, 0.113305],
     [0.13, 0.102102, 0.103397, 0.107314, 0.110731, 0.113404, 0.115477, 0.117109,
      0.118419, 0.119491, 0.120382, 0.121134, 0.121778, 0.122335, 0.122821],
     [0.14, 0.111104, 0.112066, 0.115996, 0.119529, 0.122332, 0.124524, 0.126257,
      0.127652, 0.128795, 0.129746, 0.130549, 0.131235, 0.131829, 0.132348],
     [0.15, 0.12025, 0.12086, 0.124767, 0.128387, 0.131302, 0.133602, 0.13543,
      0.136906, 0.138116, 0.139125, 0.139977, 0.140706, 0.141337, 0.141887],
     [0.16, 0.129536, 0.129782, 0.133631, 0.13731, 0.140318, 0.142714, 0.144629,
      0.146181, 0.147456, 0.14852, 0.14942, 0.15019, 0.150856, 0.151438],
     [0.17, 0.138958, 0.138834, 0.142594, 0.146304, 0.149385, 0.151863, 0.153857,
      0.155479, 0.156816, 0.157933, 0.158878, 0.159688, 0.160389, 0.161],
     [0.18, 0.148512, 0.148018, 0.151661, 0.155374, 0.158509, 0.161055, 0.163117,
      0.164803, 0.166196, 0.167363, 0.168352, 0.1692, 0.169934, 0.170575],
     [0.19, 0.158194, 0.157335, 0.160835, 0.164527, 0.167693, 0.170292, 0.172413,
      0.174154, 0.1756, 0.176813, 0.177843, 0.178727, 0.179493, 0.180162],
     [0.2, 0.168, 0.166784, 0.170121, 0.173767, 0.176944, 0.17958, 0.181747,
      0.183537, 0.185028, 0.186284, 0.187352, 0.188269, 0.189065, 0.189761],
     [0.21, 0.177926, 0.176367, 0.179523, 0.183099, 0.186268, 0.188924, 0.191125,
      0.192954, 0.194484, 0.195777, 0.19688, 0.197829, 0.198653, 0.199374],
     [0.22, 0.187968, 0.186082, 0.189043, 0.192529, 0.195669, 0.198329, 0.200552,
      0.202409, 0.203972, 0.205296, 0.206429, 0.207406, 0.208256, 0.209001],
     [0.23, 0.198122, 0.195929, 0.198684, 0.202062, 0.205153, 0.207802, 0.210032,
      0.211908, 0.213494, 0.214844, 0.216002, 0.217004, 0.217877, 0.218644],
     [0.24, 0.208384, 0.205908, 0.208449, 0.2117, 0.214726, 0.217347, 0.219571,
      0.221455, 0.223056, 0.224424, 0.225603, 0.226625, 0.227518, 0.228304],
     [0.25, 0.21875, 0.216016, 0.218338, 0.22145, 0.224394, 0.22697, 0.229176,
      0.231056, 0.232662, 0.234042, 0.235234, 0.236272, 0.237181, 0.237983],
     [0.26, 0.229216, 0.226251, 0.228354, 0.231313, 0.23416, 0.236678, 0.238852,
      0.240716, 0.242318, 0.243701, 0.244901, 0.245949, 0.24687, 0.247683],
     [0.27, 0.239778, 0.236611, 0.238496, 0.241293, 0.24403, 0.246476, 0.248604,
      0.250442, 0.25203, 0.253407, 0.254608, 0.25566, 0.256587, 0.257409],
     [0.28, 0.250432, 0.247094, 0.248764, 0.251393, 0.254008, 0.25637, 0.25844,
      0.26024, 0.261804, 0.263167, 0.264361, 0.265411, 0.266339, 0.267165],
     [0.29, 0.261174, 0.257697, 0.259158, 0.261613, 0.264097, 0.266364, 0.268365,
      0.270117, 0.271647, 0.272987, 0.274165, 0.275207, 0.276131, 0.276955],
     [0.3, 0.272, 0.268416, 0.269676, 0.271955, 0.274301, 0.276462, 0.278385,
      0.280078, 0.281564, 0.282873, 0.284029, 0.285054, 0.285968, 0.286785],
     [0.31, 0.282906, 0.279248, 0.280317, 0.28242, 0.284622, 0.28667, 0.288505,
      0.290129, 0.291564, 0.292833, 0.293958, 0.294961, 0.295857, 0.296662],
     [0.32, 0.293888, 0.290188, 0.291079, 0.293006, 0.295061, 0.296989, 0.298729,
      0.300277, 0.301652, 0.302873, 0.30396, 0.304933, 0.305806, 0.306592],
     [0.33, 0.304942, 0.301234, 0.301959, 0.303714, 0.305619, 0.307423, 0.309061,
      0.310527, 0.311834, 0.313, 0.314042, 0.314978, 0.315821, 0.316584],
     [0.34, 0.316064, 0.31238, 0.312953, 0.314541, 0.316296, 0.317973, 0.319505,
      0.320882, 0.322115, 0.32322, 0.324211, 0.325104, 0.325912, 0.326644],
     [0.35, 0.32725, 0.323621, 0.324058, 0.325485, 0.327092, 0.32864, 0.330062,
      0.331346, 0.3325, 0.333538, 0.334473, 0.335318, 0.336084, 0.336782],
     [0.36, 0.338496, 0.334954, 0.335269, 0.336542, 0.338004, 0.339423, 0.340733,
      0.341921, 0.342994, 0.34396, 0.344834, 0.345626, 0.346346, 0.347003],
     [0.37, 0.349798, 0.346372, 0.346582, 0.34771, 0.349029, 0.350321, 0.351519,
      0.35261, 0.353597, 0.35449, 0.355298, 0.356033, 0.356704, 0.357317],
     [0.38, 0.361152, 0.357871, 0.357991, 0.358982, 0.360165, 0.361332, 0.362419,
      0.363411, 0.364312, 0.365128, 0.36587, 0.366545, 0.367162, 0.367728],
     [0.39, 0.372554, 0.369445, 0.36949, 0.370355, 0.371407, 0.372452, 0.373429,
      0.374324, 0.375138, 0.375877, 0.37655, 0.377163, 0.377725, 0.378241],
     [0.4, 0.384, 0.381088, 0.381074, 0.381822, 0.382749, 0.383677, 0.384547,
      0.385346, 0.386074, 0.386736, 0.387339, 0.38789, 0.388396, 0.388861],
     [0.41, 0.395486, 0.392795, 0.392736, 0.393376, 0.394186, 0.395, 0.395767,
      0.396472, 0.397116, 0.397702, 0.398236, 0.398725, 0.399174, 0.399587],
     [0.42, 0.407008, 0.404559, 0.40447, 0.405011, 0.40571, 0.406416, 0.407084,
      0.407698, 0.408259, 0.408771, 0.409238, 0.409665, 0.410057, 0.410419],
     [0.43, 0.418562, 0.416375, 0.416268, 0.41672, 0.417313, 0.417918, 0.418489,
      0.419016, 0.419498, 0.419937, 0.420339, 0.420706, 0.421043, 0.421354],
     [0.44, 0.430144, 0.428237, 0.428122, 0.428493, 0.428988, 0.429495, 0.429975,
      0.430419, 0.430824, 0.431194, 0.431531, 0.43184, 0.432124, 0.432386],
     [0.45, 0.44175, 0.440137, 0.440026, 0.440322, 0.440726, 0.44114, 0.441533,
      0.441896, 0.442228, 0.442531, 0.442807, 0.44306, 0.443292, 0.443506],
     [0.46, 0.453376, 0.45207, 0.45197, 0.4522, 0.452516, 0.452842, 0.453151,
      0.453437, 0.453699, 0.453937, 0.454155, 0.454354, 0.454536, 0.454705],
     [0.47, 0.465018, 0.46403, 0.463949, 0.464116, 0.464349, 0.46459, 0.464819,
      0.465031, 0.465224, 0.465401, 0.465562, 0.465709, 0.465844, 0.465969],
     [0.48, 0.476672, 0.476009, 0.475952, 0.476061, 0.476214, 0.476373, 0.476525,
      0.476665, 0.476792, 0.476909, 0.477015, 0.477112, 0.477201, 0.477283],
     [0.49, 0.488334, 0.488001, 0.487972, 0.488026, 0.488102, 0.488181, 0.488256,
      0.488325, 0.488389, 0.488447, 0.4885, 0.488548, 0.488592, 0.488633],
     [0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5],
     [0.51, 0.511666, 0.511999, 0.512028, 0.511974, 0.511898, 0.511819, 0.511744,
      0.511675, 0.511611, 0.511553, 0.5115, 0.511452, 0.511408, 0.511367],
     [0.52, 0.523328, 0.523991, 0.524048, 0.523939, 0.523786, 0.523627, 0.523475,
      0.523335, 0.523208, 0.523091, 0.522985, 0.522888, 0.522799, 0.522717],
     [0.53, 0.534982, 0.53597, 0.536051, 0.535884, 0.535651, 0.53541, 0.535181,
      0.534969, 0.534776, 0.534599, 0.534438, 0.534291, 0.534156, 0.534031],
     [0.54, 0.546624, 0.54793, 0.54803, 0.5478, 0.547484, 0.547158, 0.546849,
      0.546563, 0.546301, 0.546063, 0.545845, 0.545646, 0.545464, 0.545295],
     [0.55, 0.55825, 0.559863, 0.559974, 0.559678, 0.559274, 0.55886, 0.558467,
      0.558104, 0.557772, 0.557469, 0.557193, 0.55694, 0.556708, 0.556494],
     [0.56, 0.569856, 0.571763, 0.571878, 0.571507, 0.571012, 0.570505, 0.570025,
      0.569581, 0.569176, 0.568806, 0.568469, 0.56816, 0.567876, 0.567614],
     [0.57, 0.581438, 0.583625, 0.583732, 0.58328, 0.582687, 0.582082, 0.581511,
      0.580984, 0.580502, 0.580063, 0.579661, 0.579294, 0.578957, 0.578646],
     [0.58, 0.592992, 0.595441, 0.59553, 0.594989, 0.59429, 0.593584, 0.592916,
      0.592302, 0.591741, 0.591229, 0.590762, 0.590335, 0.589943, 0.589581],
     [0.59, 0.604514, 0.607205, 0.607264, 0.606624, 0.605814, 0.605, 0.604233,
      0.603528, 0.602884, 0.602298, 0.601764, 0.601275, 0.600826, 0.600413],
     [0.6, 0.616, 0.618912, 0.618926, 0.618178, 0.617251, 0.616323, 0.615453,
      0.614654, 0.613926, 0.613264, 0.612661, 0.61211, 0.611604, 0.611139],
     [0.61, 0.627446, 0.630555, 0.63051, 0.629645, 0.628593, 0.627548, 0.626571,
      0.625676, 0.624862, 0.624123, 0.62345, 0.622837, 0.622275, 0.621759],
     [0.62, 0.638848, 0.642129, 0.642009, 0.641018, 0.639835, 0.638668, 0.637581,
      0.636589, 0.635688, 0.634872, 0.63413, 0.633455, 0.632838, 0.632272],
     [0.63, 0.650202, 0.653628, 0.653418, 0.65229, 0.650971, 0.649679, 0.648481,
      0.64739, 0.646403, 0.64551, 0.644702, 0.643967, 0.643296, 0.642683],
     [0.64, 0.661504, 0.665046, 0.664731, 0.663458, 0.661996, 0.660577, 0.659267,
      0.658079, 0.657006, 0.65604, 0.655166, 0.654374, 0.653654, 0.652997],
     [0.65, 0.67275, 0.676379, 0.675942, 0.674515, 0.672908, 0.67136, 0.669938,
      0.668654, 0.6675, 0.666462, 0.665527, 0.664682, 0.663916, 0.663218],
     [0.66, 0.683936, 0.68762, 0.687047, 0.685459, 0.683704, 0.682027, 0.680495,
      0.679118, 0.677885, 0.67678, 0.675789, 0.674896, 0.674088, 0.673356],
     [0.67, 0.695058, 0.698766, 0.698041, 0.696286, 0.694381, 0.692577, 0.690939,
      0.689473, 0.688166, 0.687, 0.685958, 0.685022, 0.684179, 0.683416],
     [0.68, 0.706112, 0.709812, 0.708921, 0.706994, 0.704939, 0.703011, 0.701271,
      0.699723, 0.698348, 0.697127, 0.69604, 0.695067, 0.694194, 0.693408],
     [0.69, 0.717094, 0.720752, 0.719683, 0.71758, 0.715378, 0.71333, 0.711495,
      0.709871, 0.708436, 0.707167, 0.706042, 0.705039, 0.704143, 0.703338],
     [0.7, 0.728, 0.731584, 0.730324, 0.728045, 0.725699, 0.723538, 0.721615,
      0.719922, 0.718436, 0.717127, 0.715971, 0.714946, 0.714032, 0.713215],
     [0.71, 0.738826, 0.742303, 0.740842, 0.738387, 0.735903, 0.733636, 0.731635,
      0.729883, 0.728353, 0.727013, 0.725835, 0.724793, 0.723869, 0.723045],
     [0.72, 0.749568, 0.752906, 0.751236, 0.748607, 0.745992, 0.74363, 0.74156,
      0.73976, 0.738196, 0.736833, 0.735639, 0.734589, 0.733661, 0.732835],
     [0.73, 0.760222, 0.763389, 0.761504, 0.758707, 0.75597, 0.753524, 0.751396,
      0.749558, 0.74797, 0.746593, 0.745392, 0.74434, 0.743413, 0.742591],
     [0.74, 0.770784, 0.773749, 0.771646, 0.768687, 0.76584, 0.763322, 0.761148,
      0.759284, 0.757682, 0.756299, 0.755099, 0.754051, 0.75313, 0.752317],
     [0.75, 0.78125, 0.783984, 0.781662, 0.77855, 0.775606, 0.77303, 0.770824,
      0.768944, 0.767338, 0.765958, 0.764766, 0.763728, 0.762819, 0.762017],
     [0.76, 0.791616, 0.794092, 0.791551, 0.7883, 0.785274, 0.782653, 0.780429,
      0.778545, 0.776944, 0.775576, 0.774397, 0.773375, 0.772482, 0.771696],
     [0.77, 0.801878, 0.804071, 0.801316, 0.797938, 0.794847, 0.792198, 0.789968,
      0.788092, 0.786506, 0.785156, 0.783998, 0.782996, 0.782123, 0.781356],
     [0.78, 0.812032, 0.813918, 0.810957, 0.807471, 0.804331, 0.801671, 0.799448,
      0.797591, 0.796028, 0.794704, 0.793571, 0.792594, 0.791744, 0.790999],
     [0.79, 0.822074, 0.823633, 0.820477, 0.816901, 0.813732, 0.811076, 0.808875,
      0.807046, 0.805516, 0.804223, 0.80312, 0.802171, 0.801347, 0.800626],
     [0.8, 0.832, 0.833216, 0.829879, 0.826233, 0.823056, 0.82042, 0.818253,
      0.816463, 0.814972, 0.813716, 0.812648, 0.811731, 0.810935, 0.810239],
     [0.81, 0.841806, 0.842665, 0.839165, 0.835473, 0.832307, 0.829708, 0.827587,
      0.825846, 0.8244, 0.823187, 0.822157, 0.821273, 0.820507, 0.819838],
     [0.82, 0.851488, 0.851982, 0.848339, 0.844626, 0.841491, 0.838945, 0.836883,
      0.835197, 0.833804, 0.832637, 0.831648, 0.8308, 0.830066, 0.829425],
     [0.83, 0.861042, 0.861166, 0.857406, 0.853696, 0.850615, 0.848137, 0.846143,
      0.844521, 0.843184, 0.842067, 0.841122, 0.840312, 0.839611, 0.839],
     [0.84, 0.870464, 0.870218, 0.866369, 0.86269, 0.859682, 0.857286, 0.855371,
      0.853819, 0.852544, 0.85148, 0.85058, 0.84981, 0.849144, 0.848562],
     [0.85, 0.87975, 0.87914, 0.875233, 0.871613, 0.868698, 0.866398, 0.86457,
      0.863094, 0.861884, 0.860875, 0.860023, 0.859294, 0.858663, 0.858113],
     [0.86, 0.888896, 0.887934, 0.884004, 0.880471, 0.877668, 0.875476, 0.873743,
      0.872348, 0.871205, 0.870254, 0.869451, 0.868765, 0.868171, 0.867652],
     [0.87, 0.897898, 0.896603, 0.892686, 0.889269, 0.886596, 0.884523, 0.882891,
      0.881581, 0.880509, 0.879618, 0.878866, 0.878222, 0.877665, 0.877179],
     [0.88, 0.906752, 0.905149, 0.901285, 0.898011, 0.895486, 0.893541, 0.892016,
      0.890795, 0.889796, 0.888966, 0.888266, 0.887666, 0.887148, 0.886695],
     [0.89, 0.915454, 0.913576, 0.909807, 0.906703, 0.904341, 0.902533, 0.90112,
      0.89999, 0.899067, 0.8983, 0.897652, 0.897098, 0.896618, 0.896199],
     [0.9, 0.924, 0.921888, 0.918257, 0.91535, 0.913163, 0.9115, 0.910203,
      0.909167, 0.908321, 0.907618, 0.907024, 0.906516, 0.906076, 0.905692],
     [0.91, 0.932386, 0.93009, 0.926642, 0.923955, 0.921957, 0.920444, 0.919267,
      0.918327, 0.917559, 0.916921, 0.916382, 0.915921, 0.915522, 0.915174],
     [0.92, 0.940608, 0.938186, 0.934967, 0.932522, 0.930723, 0.929366, 0.928311,
      0.927469, 0.926782, 0.92621, 0.925727, 0.925314, 0.924956, 0.924644],
     [0.93, 0.948662, 0.946183, 0.943237, 0.941055, 0.939463, 0.938266, 0.937336,
      0.936594, 0.935988, 0.935484, 0.935058, 0.934694, 0.934378, 0.934103],
     [0.94, 0.956544, 0.954088, 0.951458, 0.949556, 0.948179, 0.947145, 0.946343,
      0.945702, 0.945179, 0.944744, 0.944376, 0.944061, 0.943789, 0.94355],
     [0.95, 0.96425, 0.961906, 0.959635, 0.958028, 0.956871, 0.956004, 0.955331,
      0.954794, 0.954355, 0.953989, 0.95368, 0.953416, 0.953187, 0.952986],
     [0.96, 0.971776, 0.969646, 0.967773, 0.966473, 0.965541, 0.964843, 0.964301,
      0.963868, 0.963514, 0.96322, 0.962971, 0.962758, 0.962573, 0.962411],
     [0.97, 0.979118, 0.977316, 0.975875, 0.974892, 0.974189, 0.973662, 0.973253,
      0.972926, 0.972659, 0.972436, 0.972248, 0.972087, 0.971947, 0.971825],
     [0.98, 0.986272, 0.984925, 0.983946, 0.983285, 0.982814, 0.982461, 0.982187,
      0.981967, 0.981788, 0.981638, 0.981512, 0.981404, 0.98131, 0.981228],
     [0.99, 0.993234, 0.992483, 0.991987, 0.991655, 0.991418, 0.99124, 0.991102,
      0.990992, 0.990902, 0.990826, 0.990763, 0.990708, 0.990661, 0.990619],
     [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]]

def set_preliminary_felo_ratings(fencers, bout, parameters):
    """Calculates the new Felo numbers of the two fencers of a given bout and
    stores them at preliminary places in the fencer objects.  More accurately,
    the new Felo numbers are stored in the felo_rating_preliminary attribute of
    the fencers.  The reason is that some bouts may take place on the same day,
    without knowing the order.  Then only preliminary values are determined and
    merged with the real ones when all bouts of that day have been considered.

    :Parameters:
      - `fencers`: all fencers.  The two fencers in it which take part in the
        bout are modified.
      - `bout`: the bout that is to be calculated.
      - `parameters`: all Felo parameters.

    :type fencers: dict
    :type bout: Bout
    :type parameters: dict
    """
    first_fencer, second_fencer, points_first, points_second = \
        bout.first_fencer, bout.second_fencer, \
        bout.points_first, bout.points_second
    max_points = max(points_first, points_second)
    total_points = points_first + points_second
    if bout.fenced_to == 0:
        weighting = parameters["weighting team bout"]
    else:
        # weighting roughly is the number of bouts fenced to 5 points
        weighting = total_points / 6.76
    if total_points == 0:
        result_first = 0.5
    else:
        result_first = float(points_first) / total_points
    if fencers[first_fencer].freshman and fencers[second_fencer].freshman:
        # Two freshmen, so the bout cannot be counted at all
        return
    if fencers[first_fencer].freshman:
        fencers[first_fencer].total_weighting += weighting
        fencers[first_fencer].total_result += (result_first - 0.5) * weighting
        fencers[first_fencer].total_felo_rating_opponents += \
            fencers[second_fencer].felo_rating_exact * weighting
        return
    elif fencers[second_fencer].freshman:
        fencers[second_fencer].total_weighting += weighting
        fencers[second_fencer].total_result += (0.5 - result_first) * weighting
        fencers[second_fencer].total_felo_rating_opponents += \
            fencers[first_fencer].felo_rating_exact * weighting
        return
    # Use current (rather than preliminary) numbers for the calculation.  Just
    # don't *store* the results in the real attributes.
    felo_first = fencers[first_fencer].felo_rating_exact
    felo_second = fencers[second_fencer].felo_rating_exact
    expectation_first = 1 / (1 + 10**((felo_second - felo_first)/400.0))
    if bout.fenced_to == 0 and max_points <= 15 and expectation_first < 1.0:
        # Adjusting the expectation value to eliminate the bias due to the
        # winning-hit problem.  I interpolate between two adjactent points in
        # the value array apparent_expectation_values.  Interpolating is
        # necessary, otherwise the bootstrapping doesn't converge.  Even with
        # this simple linear interpolating, convergence is significantly more
        # difficult than without the winning-hit adjustment at all.  This
        # adjustment is only necessary if a bout is not weighted according to
        # the total points fenced.  At the moment, this is only the case for
        # single bouts in a team relay competition.
        expectation_first = (apparent_expectation_values[int(expectation_first*100)+1][max_points-1] -
                             apparent_expectation_values[int(expectation_first*100)][max_points-1]) * \
                             (expectation_first*100 - int(expectation_first*100)) + \
                             apparent_expectation_values[int(expectation_first*100)][max_points-1]
    improvement_first = (result_first - expectation_first) * weighting
    fencers[first_fencer].felo_rating_preliminary += fencers[first_fencer].k_factor * improvement_first
    fencers[second_fencer].felo_rating_preliminary -= fencers[second_fencer].k_factor * improvement_first
    Fencer.fencers_with_preliminary_felo_rating.add(fencers[first_fencer])
    Fencer.fencers_with_preliminary_felo_rating.add(fencers[second_fencer])

    fencers[first_fencer].total_weighting_preliminary += weighting
    fencers[second_fencer].total_weighting_preliminary += weighting

def parse_bouts(input_file, linenumber, fencers, parameters):
    """Reads bouts from a Felo file, starting at the current position in that file.
    It reads till the end of the file since the bouts are the last section in a
    Felo file.

    :Parameters:
      - `input_file`: an *open* file object where the items are read from
      - `linenumber`: number of already read lines in the input_file
      - `fencers`: all fencers.  Foreign fencers may be added to it
      - `parameters`: all Felo parameters

    :type input_file: file
    :type linenumber: int
    :type fencers: dict
    :type parameters: dict

    :Return:
      - list with all bouts that were read.

    :rtype: list

    :Exceptions:
      - `LineError`: if a line does not follow the Felo file syntax
    """
    line_pattern = re.compile("\\s*(?:(?:(?P<year>\\d{4})-(?P<month>\\d{1,2})-(?P<day>\\d{1,2}))?"
                              "(?:\\.?(?P<index>\\d+))?"+separator+")?"+
                              "(?P<first>.+?)\\s*--\\s*(?P<second>.+?)"+separator+
                              "(?P<points_first>\\d+):(?P<points_second>\\d+)\\s*"+
                              "(?P<fenced_to>(?:/\\d+)|\\*)?\\s*\\Z")
    bouts = []
    for line in input_file:
        linenumber += 1
        line = clean_up_line(line)
        if not line: continue
        match = line_pattern.match(line)
        if not match:
            raise LineError(_('Line must follow the pattern "YYYY-MM-DD <TAB> name1 '
                              '-- name2 <TAB> points1:points2".'),
                        input_file.name, linenumber)
        year, month, day, index, first_fencer, second_fencer, points_first, points_second, fenced_to = \
            match.groups()
        points_first, points_second = int(points_first), int(points_second)
        if not fenced_to:
            fenced_to = max(points_first, points_second)
        else:
            if fenced_to == "*":
                fenced_to = "/0"
            fenced_to = int(fenced_to[1:])
        if fenced_to > 0 and (points_first > fenced_to or points_second > fenced_to):
            raise LineError(_("One fencer has more points than the winning points."), input_file.name, linenumber)
        if first_fencer not in fencers:
            if first_fencer.find("<") == -1:
                raise LineError(_('Fencer "%s" is unknown.') % first_fencer, input_file.name, linenumber)
            # The -1 is irrelevant; it is set in the constructor anyway
            fencers[first_fencer] = Fencer(first_fencer, -1, parameters)
        if second_fencer not in fencers:
            if second_fencer.find("<") == -1:
                raise LineError(_('Fencer "%s" is unknown.') % second_fencer, input_file.name, linenumber)
            # The -1 is irrelevant; it is set in the constructor anyway
            fencers[second_fencer] = Fencer(second_fencer, -1, parameters)
        if not index:
            index = "0"
        if year:
            last_year, last_month, last_day = year, month, day
        else:
            try:
                year, month, day = last_year, last_month, last_day
            except NameError:
                raise LineError(_('No date found for this bout.'), input_file.name, linenumber)
        bouts.append(Bout(int(year), int(month), int(day), int(index), first_fencer, second_fencer,
                                points_first, points_second, fenced_to))
    return bouts

def parse_felo_file(felo_file):
    """Reads from a Felo file parameters, fencers, and bouts.

    :Parameters:
      - `felo_file`: the *open* file to be parsed

    :type felo_file: file

    :Return:
      - Felo parameters as a dictionary
      - a list with all really in the file given parameters
      - fencers as a dictionary (in the form "name: Fencer object")
      - bouts as a list.

    :rtype: dict, list, dict, list

    :Exceptions:
      - `FeloFormatError`: if an initial Felo rating or weighting is invalid
    """
    english_parameter_names = {_(u"k factor top fencers"): "k factor top fencers",
                               _(u"felo rating top fencers"): "felo rating top fencers",
                               _(u"k factor others"): "k factor others",
                               _(u"k factor freshmen"): "k factor freshmen",
                               _(u"5 point bouts freshmen"): "5 point bouts freshmen",
                               _(u"minimal felo rating"): "minimal felo rating",
                               _(u"5 point bouts for estimate"): "5 point bouts for estimate",
                               _(u"groupname"): "groupname",
                               _(u"output folder"): "output folder",
                               _(u"min distance of plot tics"): "min distance of plot tics",
                               _(u"earliest date in plot"): "earliest date in plot",
                               _(u"maximal days in plot"): "maximal days in plot",
                               _(u"threshold bootstrapping"): "threshold bootstrapping",
                               _(u"weighting team bout"): "weighting team bout",
                               _(u"path of gnuplot"): "path of gnuplot",
                               _(u"path of convert"): "path of convert",
                               _(u"path of ps2pdf"): "path of ps2pdf"}
    parameters_native_language, linenumber = parse_items(felo_file)
    parameters = {}
    for native_name, value in parameters_native_language.items():
        if native_name in english_parameter_names:
            parameters[english_parameter_names[native_name]] = value
        elif native_name in english_parameter_names.values():
            # The non-English user used an English parameter name
            parameters[native_name] = value
        else:
            raise FeloFormatError(_(u"Parameter \"%s\" is unknown.") % native_name)
    given_parameters = list(parameters)
    parameters.setdefault("k factor top fencers", 25)
    parameters.setdefault("felo rating top fencers", 2400)
    parameters.setdefault("k factor others", 32)
    parameters.setdefault("k factor freshmen", 40)
    parameters.setdefault("5 point bouts freshmen", 15)
    parameters.setdefault("minimal felo rating", 1200)
    parameters.setdefault("5 point bouts for estimate", 10)
    parameters.setdefault("threshold bootstrapping", 0.001)
    parameters.setdefault("weighting team bout", 1.0)
    # The groupname is used e.g. for the file names of the plots.  It defaults
    # to the name of the Felo file.
    parameters.setdefault("groupname",
                          os.path.splitext(os.path.split(felo_file.name)[1])[0].capitalize())
    parameters.setdefault("output folder", os.path.abspath(os.path.dirname(felo_file.name)))
    parameters.setdefault("min distance of plot tics", 7)
    parameters.setdefault("earliest date in plot", "1980-01-01")
    parameters.setdefault("maximal days in plot", "366")
    if os.name == 'nt':
        programs_path = os.environ["ProgramFiles"]
        parameters.setdefault("path of gnuplot", os.path.join(programs_path, r"gnuplot\bin\wgnuplot.exe"))
        # ImageMagick and Ghostscript include a version number into their
        # paths, thus I have to do some extra work.
        possible_paths = glob.glob(os.path.join(programs_path, "ImageMagick*"))
        if possible_paths:
            parameters.setdefault("path of convert", os.path.join(possible_paths[0], "convert.exe"))
        else:
            parameters.setdefault("path of convert", "convert")
        # Now Ghostscript
        possible_paths = glob.glob(os.path.join(programs_path, "gs", "gs*"))
        if possible_paths:
            parameters.setdefault("path of ps2pdf", os.path.join(possible_paths[0], "lib", "ps2pdf.bat"))
        else:
            parameters.setdefault("path of ps2pdf", "ps2pdf")
    else:
        parameters.setdefault("path of gnuplot", "gnuplot")
        parameters.setdefault("path of convert", "convert")
        parameters.setdefault("path of ps2pdf", "ps2pdf")

    initial_felo_ratings, linenumber = parse_items(felo_file, linenumber)
    fencers = {}
    for name, initial_felo_rating in initial_felo_ratings.items():
        initial_total_weighting = 0.0
        if isinstance(initial_felo_rating, str):
            try:
                # An initial total weighting is also available, in parenthesis
                position_opening_parenthesis = initial_felo_rating.index("(")
                initial_total_weighting = float(initial_felo_rating[position_opening_parenthesis+1:-1])
                initial_felo_rating = int(initial_felo_rating[:position_opening_parenthesis])
            except ValueError:
                raise FeloFormatError(_(u"Felo rating of \"%s\" is invalid.") % name)
        current_fencer = Fencer(name, initial_felo_rating, parameters, initial_total_weighting)
        fencers[current_fencer.name] = current_fencer

    bouts = parse_bouts(felo_file, linenumber, fencers, parameters)
    return parameters, given_parameters, fencers, bouts

def fill_with_tabs(text, tab_col):
    """Adds tabs to a string until a certain column is reached.

    :Parameters:
      - `text`: the string that should be expanded to a certain column
      - `tab_col`: the tabulator column to which should be filled up.  It is *not*
        a character column, so a value of, say, "4" means character column 32.

    :type text: string
    :type tab_col: int

    :Return:
      - The expanded string, i.e., "text" plus one or more tabs.

    :rtype: string
    """
    number_of_tabs = tab_col - len(text.expandtabs()) / 8
    return text + max(number_of_tabs, 1) * "\t"

def write_felo_file(filename, parameters, fencers, bouts):
    """Write Felo parameters, fencers, and bouts to a Felo file, which is
    overwritten if already existing.

    :Parameters:
      - `filename`: path to the Felo file to be written
      - `parameters`: Felo parameters
      - `fencers`: all fencers
      - `bouts`: all bouts

    :type filename: string
    :type parameters: dict
    :type fencers: dict
    :type bouts: list
    """
    felo_file = codecs.open(filename, "w", "utf-8")
    print(_("# Parameters"), file=felo_file)
    print>>felo_file
    parameterslist = parameters.items()
    # sort case-insensitively
    parameterslist.sort(lambda x,y: x[0].upper() < y[0].upper())
    for name, value in parameterslist:
        print>>felo_file, fill_with_tabs(name, 4) + str(value)
    print>>felo_file
    print>>felo_file, 52 * "="
    print>>felo_file, _("# Initial Felo ratings")
    print>>felo_file, _("# Names of fencers who want to be hidden")
    print>>felo_file, _("# in parentheses")
    print>>felo_file
    fencerslist = fencers.items()
    fencerslist.sort()
    for fencer in [entry[1] for entry in fencerslist]:
        if fencer.hidden:
            name = "(" + fencer.name + ")"
        else:
            name = fencer.name
        print>>felo_file, fill_with_tabs(name, 3) + str(fencer.initial_felo_rating)
    print>>felo_file
    print>>felo_file, 52 * "="
    print>>felo_file, _("# Bouts")
    print>>felo_file
    bouts.sort()
    for i, bout in enumerate(bouts):
        # Separate days for clarity
        if i > 0 and bouts[i-1].date.toordinal() != bout.date.toordinal():
            print>>felo_file
        line = fill_with_tabs(fill_with_tabs(bout.date_string, 2) +
                               "%s -- %s" % (bout.first_fencer, bout.second_fencer), 5) + \
                               "%d:%d" % (bout.points_first, bout.points_second)
        if bout.fenced_to == 0:
            line += " *"
        elif bout.points_first != bout.fenced_to and bout.points_second != bout.fenced_to:
            line += "/" + str(bout.fenced_to)
        print>>felo_file, line
    felo_file.close()

def write_back_fencers(felo_file_contents, fencers):
    """Write *only* the fencers section back into a Felo file, given as a string
    instead of a file.  The rest of the Felo file remains untouched.  Moreover,
    all introducing comments and empty lines in the fencers section remain
    untouched, too.

    :Parameters:
      - `felo_file_contents`: the old version of the Felo file
      - `fencers`: all fencers, which replace the fencers section in the Felo
        file

    :type felo_file_contents: string
    :type fencers: dict

    :Return:
      - the new Felo file as a string

    :rtype: string

    :Exceptions:
      - `FeloFormatError`: if the original Felo file is invalid concerning the
        so-called "boundaries" (the "=====" lines)
    """
    lines = felo_file_contents.splitlines()
    fencer_limits = []
    for linenumber, line in enumerate(lines):
        cleaned_line = clean_up_line(line)
        if cleaned_line and cleaned_line[0] in u"-=._:;,+*'~\"`´/\\%$!":
            fencer_limits.append(linenumber)
    if len(fencer_limits) != 2:
        raise FeloFormatError(_("Felo file invalid because there are not exactly two boundary lines."))
    fencer_lines = []
    while lines[fencer_limits[0] + 1].lstrip().startswith("#") or lines[fencer_limits[0] + 1].lstrip() == "":
        fencer_limits[0] += 1
    fencerslist = fencers.items()
    fencerslist.sort()
    for fencer in [entry[1] for entry in fencerslist]:
        if fencer.hidden:
            name = "(" + fencer.name + ")"
        else:
            name = fencer.name
        line = fill_with_tabs(name, 3) + str(fencer.initial_felo_rating)
        if fencer.initial_total_weighting != 0:
            line += " (%g)" % fencer.initial_total_weighting
        fencer_lines.append(line)
    fencer_lines.append(u"")
    return "\n".join(lines[:fencer_limits[0]+1] + fencer_lines + lines[fencer_limits[1]:]) + "\n"

def write_back_fencers_to_file(filename, fencers):
    """Write *only* the fencers section back into a Felo file.  The rest of the
    Felo file remains untouched.  Moreover, all introducing comments and empty
    lines in the fencers section remain untouched, too.

    :Parameters:
      - `filename`: path to the old version of the Felo file
      - `fencers`: all fencers, which replace the fencers section in the Felo
        file

    :type filename: string
    :type fencers: dict

    :Exceptions:
      - `FeloFormatError`: if the original Felo file is invalid concerning the
        so-called "boundaries" (the "=====" lines)
    """
    filename_backup = os.path.splitext(filename)[0] + ".bak"
    if os.path.isfile(filename_backup):
        raise Error(_(u"First delete the backup (.bak) file."))
    shutil.copyfile(filename, filename_backup)
    felo_file = codecs.open(filename, encoding="utf-8")
    contents = felo_file.read()
    felo_file.close()
    felo_file = codecs.open(filename, "w", encoding="utf-8")
    felo_file.write(write_back_fencers(contents, fencers))
    felo_file.close()

class Error(Exception):
    """Standard error class.
    """
    def __init__(self, description):
        """Class constructor.

        :Parameters:
          - `description`: error message

        :type description: string
        """
        self.description = description
        if isinstance(description, str):
            # FixMe: The following must be made OS-dependent
            description = description.encode("utf-8")
        Exception.__init__(self, description)

class LineError(Error):
    """Error class for parsing errors in a Felo file.
    """
    def __init__(self, description, filename, linenumber):
        """Class constructor.

        :Parameters:
          - `description`: error message
          - `filename`: Felo file name
          - `linenumber`: number of the line in the Felo file where the error
            occured

        :type description: string
        :type filename: string
        :type linenumber: int
        """
        self.linenumber = linenumber
        self.naked_description = description
        if filename:
            supplement = filename
            if linenumber:
                supplement += _(u", line ") + str(linenumber)
            description = supplement  + ": " + description
        Error.__init__(self, description)

class FeloFormatError(Error):
    """Error class for invalid values in the Felo file.
    """
    def __init__(self, description):
        """Class constructor.

        :Parameters:
          - `description`: error message

        :type description: string
        """
        Error.__init__(self, description)

class BootstrappingError(Error):
    """Error class for non-converging bootstrapping processes.
    """
    def __init__(self, description):
        """Class constructor.

        :Parameters:
          - `description`: error message

        :type description: string
        """
        Error.__init__(self, description)

class ExternalProgramError(Error):
    """Error class for external programs that were not found.
    """
    def __init__(self, description):
        """Class constructor.

        :Parameters:
          - `description`: error message

        :type description: string
        """
        Error.__init__(self, description)

def adopt_preliminary_felo_ratings():
    """Take all preliminary numbers (Felo and fenced points) and write them over
    the "real" numbers.  The preliminary numbers remain untouched and can be
    used for further calculations.

    For optimisation, only fencers in the set
    fencers_with_preliminary_felo_rating, which is a class variable in Fencers,
    are visited.

    No Parameters and return values.
    """
    for fencer in Fencer.fencers_with_preliminary_felo_rating:
        fencer.felo_rating = fencer.felo_rating_preliminary
        fencer.total_weighting = fencer.total_weighting_preliminary
    Fencer.fencers_with_preliminary_felo_rating.clear()


class Fencer(object):
    """Class for fencer data.  Basically, it is a mere container for the
    attributes.

    :cvar fencers_with_preliminary_felo_rating: all fencers which still have
      their Felo number in felo_rating_preliminary, so that it must be copied to
      felo_rating is a bout day is completely processed.

    :type fencers_with_preliminary_felo_rating: set
    """
    fencers_with_preliminary_felo_rating = set()

    def __init__(self, name, felo_rating, parameters, initial_total_weighting=0):
        """Class constructor.

        :Parameters:
          - `name`: name of the fencer
          - `felo_rating`: initial Felo rating
          - `parameters`: all Felo parameters

        :type name: string
        :type felo_rating: int
        :type parameters: dict
        """
        self.hidden = name.startswith(u"(")
        if self.hidden:
            self.name = name[1:-1]
        else:
            self.name = name
        self.parameters = parameters
        # total_weighting is the number of bouts that the fencers has fenced,
        # in units of "bout to 5 points"-equivalents.
        self.total_weighting = self.initial_total_weighting = self.total_weighting_preliminary = \
            initial_total_weighting
        self.__k_factor = self.parameters["k factor others"]
        self.foreign_fencer = name.find("<") != -1
        if self.foreign_fencer:
            try:
                self.__felo_rating = int(re.search(r"<(\d+)>", name).group(1))
                if self.__felo_rating <= 0:
                    raise ValueError
            except (ValueError, AttributeError):
                raise Error("Foreign fencer '%s' has invalid Felo rating" % name)
        self.freshman = felo_rating == 0
        if not self.freshman:
            self.felo_rating = self.initial_felo_rating = self.felo_rating_preliminary = felo_rating
        else:
            self.total_felo_rating_opponents = 0.0
            self.initial_felo_rating = 0
            self.total_result = 0.0

    def __get_felo_rating_exact(self):
        if not self.freshman:
            return self.__felo_rating
        else:
            # Estimate initial Felo number according to the Austrian Method,
            # see http://www.chess.at/bundesspielleitung/OESB/oesb_tuwo_06.pdf
            # section 5.1 on page 43.
            if self.total_weighting < self.parameters["5 point bouts for estimate"] or \
                            self.total_weighting == 0:
                return 0.0
            A = self.total_result / self.total_weighting
            B = self.total_weighting / (self.total_weighting + 2)
            average_felo_rating_opponents = self.total_felo_rating_opponents / self.total_weighting
            return average_felo_rating_opponents + (A * B * 700)

    def __get_felo_rating(self):
        return int(round(self.felo_rating_exact))

    def __set_felo_rating(self, felo_rating):
        if not self.freshman and not self.foreign_fencer:
            self.__felo_rating = max(felo_rating, self.parameters["minimal felo rating"])
            if self.__felo_rating >= self.parameters["felo rating top fencers"]:
                self.__k_factor = self.parameters["k factor top fencers"]

    felo_rating_exact = property(__get_felo_rating_exact, __set_felo_rating,
                                 doc="""Felo rating with decimal fraction.
                                        @type: float""")
    felo_rating = property(__get_felo_rating, __set_felo_rating,
                           doc="""Felo rating, rounded to integer.
                                  @type: int""")

    def __get_k_factor(self):
        if self.total_weighting < self.parameters["5 point bouts freshmen"]:
            return self.parameters["k factor freshmen"]
        return self.__k_factor

    k_factor = property(__get_k_factor, doc="""k factor (see Elo formula) of this fencer.
                                               @type: int""")

    def __cmp__(self, other):
        """Sort by Felo rating, descending."""
        return -(self.felo_rating_exact - other.felo_rating_exact)

    def __repr__(self):
        """No formal representation available yet, thus I call __str__()."""
        return self.__str__()

    def __str__(self):
        """Informal string representation of the fencer."""
        return self.name + " (" + str(self.felo_rating) + ")"


def calculate_felo_ratings(parameters, fencers, bouts, plot=False, estimate_freshmen=False,
                           bootstrapping=False, maxcycles=1000, bootstrapping_callback=None):
    """Calculate the new Felo ratings, taking a whole bunch of bouts into
    account.  If wanted, generate plots with the development of the Felo
    numbers.

    :Parameters:
      - `parameters`: all Felo parameters
      - `fencers`: all fencers
      - `bouts`: all bouts, which is sorted chronologically by this function if
        necessary
      - `plot`: if True, plots as PNG and PDF are generated.  The file name is the
        group name in the Felo parameters.
      - `bootstrapping`: if True, try to estimate good starting Felo numbers by
        going through the bouts multiple times, using the result numbers as new
        starting numbers, until the numbers have converged.
      - `maxcycles`: the maximal number of cycles in the bootstrapping process.
        If it is not enough, an exception is raised.
      - `estimate_freshmen`: if True, try to calculate estimates for freshmen.
      - `bootstrapping_callback`: callabale object which takes as the only
        parameter a float between 0 and 1 indicating the progress while
        bootstrapping.  It may update a progress bar, for example.

    :type parameters: dict
    :type fencers: dict
    :type bouts: list
    :type plot: boolean
    :type bootstrapping: boolean
    :type maxcycles: int
    :type estimate_freshmen: boolean
    :type bootstrapping_callback: callable

    :Return:
      - list (not a dictionary!) of all visible fencers, sorted by descending
        Felo number
      - If C{estimate_freshmen} is True, a list with all freshmen

    :rtype: list, list

    :Exceptions:
      - `BootstrappingError`: if the bootstrapping didn't converge
      - `ExternalProgramError`: if an external program is not found
    """

    def calculate_felo_ratings_core(parameters, fencers, bouts, plot, data_file_name=None):
        """Calculate the new Felo ratings, taking a whole bunch of bouts into
        account.  If wanted, generate plots with the development of the Felo
        numbers.  Some things that are necessary before and after are not done
        here, because it is only the "core" routine.

        :Parameters:
          - `parameters`: dictionary with all Felo parameters
          - `fencers`: dictionary with all fencers
          - `bouts`: list of bouts, which is sorted chronologically if necessary
          - `plot`: If True, plots as PNG and PDF are generated.  The file name is the
            group name in the Felo parameters.
          - `data_file_name`: just a copy of the variable of the same name
            from the enclosing scope.  It's the filename of the datafile
            plotted by Gnuplot.

        :type parameters: dict
        :type fencers: dict
        :type bouts: list
        :type plot: boolean
        :type data_file_name: string

        :Return:
          - the accumulated xtics, but only if C{plot==True}

        :rtype: string
        """
        if plot:
            data_file = open(data_file_name, "w")
            # xtics holds the Gnuplot command for the x axis labels (i.e. the dates).
            xtics = ""
            last_xtics_daynumber = 0
        for i, bout in enumerate(bouts):
            set_preliminary_felo_ratings(fencers, bout, parameters)
            if i == len(bouts) - 1 or bout.date_string != bouts[i + 1].date_string:
                # Not one *day* is over but one set of bouts which took place with
                # unknown order.
                adopt_preliminary_felo_ratings()
            current_bout_daynumber = bout.date.toordinal()
            year, month, day, __, __, __, __, __, __ = time.localtime()
            current_daynumber = datetime.date(year, month, day).toordinal()
            # There are three conditions so that plot points are created: We
            # must have a new day, so we don't generate points within one day
            # (too fine-grained, and no real times are known within one day
            # anyway); the Felo ratings must be after the minimal date given in
            # the parameters section; *and* the Felo ratings must not be older
            # than the maximal days in the plot.
            if plot and (i == len(bouts) - 1 or bouts[i + 1].date.toordinal() != current_bout_daynumber) and \
                            bout.date_string[:10] >= parameters["earliest date in plot"] and \
                                    current_daynumber - current_bout_daynumber <= parameters["maximal days in plot"]:
                data_file.write(str(current_bout_daynumber))
                if current_bout_daynumber - last_xtics_daynumber >= parameters["min distance of plot tics"]:
                    # Generate tic marks not too densely; labels must be at
                    # least the the minimal tic distance apart.
                    last_xtics_daynumber = current_bout_daynumber
                    xtics += bout.date.strftime(str(_(u"'%Y-%m-%d'"))) + " %d," % current_bout_daynumber
                for fencer in visible_fencers:
                    data_file.write("\t" + str(fencer.felo_rating_exact))
                data_file.write(os.linesep)
        if plot:
            data_file.close()
            return xtics

    def construct_supplement(path):
        """Builds a message with the path where felo_rating.py looked for an
        external program but didn't find it.

        :Parameters:
          - `path`: path where it was looked for the program

        :type path: string

        :Return:
          - the error message supplement as a string

        :rtype: string
        """
        if os.name == 'nt':
            return _(u"(I looked for it at '%s'.)  ") % path
        else:
            return ""

    bouts_base_filename = parameters["groupname"].lower()
    tempdir = tempfile.gettempdir()
    data_file_name = os.path.join(tempdir, bouts_base_filename + ".dat")
    gnuplot_script_file_name = os.path.join(tempdir, bouts_base_filename + ".gp")
    postscript_file_name = os.path.join(tempdir, bouts_base_filename + ".ps")
    png_file_name = os.path.join(parameters["output folder"], bouts_base_filename + ".png")
    pdf_file_name = os.path.join(parameters["output folder"], bouts_base_filename + ".pdf")

    visible_fencers = [fencer for fencer in fencers.values()
                       if not (fencer.hidden or fencer.freshman or fencer.foreign_fencer)]
    for index, fencer in enumerate(visible_fencers):
        # Store the column index of the data file in the fencer object.  Needed
        # by Gnuplot.
        fencer.columnindex = index + 2
    bouts.sort()
    if bootstrapping:
        for i in range(maxcycles):
            if i % 10 == 0 and bootstrapping_callback:
                bootstrapping_callback(float(i) / (maxcycles - 1))
            for fencer in fencers.values():
                fencer.old_felo_rating = fencer.felo_rating_exact
            calculate_felo_ratings_core(parameters, fencers, bouts, plot=False)
            for fencer in fencers.values():
                if abs(fencer.old_felo_rating - fencer.felo_rating_exact) >= parameters["threshold bootstrapping"]:
                    break
            else:
                break
        if i == maxcycles - 1:
            raise BootstrappingError(_("The bootstrapping didn't converge."))
    xtics = calculate_felo_ratings_core(parameters, fencers, bouts, plot, data_file_name)
    visible_fencers.sort()  # Descending by Felo rating
    if plot:
        # Call Gnuplot, convert, and ps2pdf to generate the PNG and PDF plots.
        # Note: We don't generate HTML tables here.  These must be provided
        # separately.
        gnuplot_script = codecs.open(gnuplot_script_file_name, "w", encoding="utf-8")
        gnuplot_script.write(u"set term postscript color; set output '" + postscript_file_name + "';"
                                                                                                 u"set key outside; set xtics rotate; set grid xtics;"
                                                                                                 u"set xtics nomirror (%s);" % xtics[
                                                                                                                               :-1] +
                             u"plot ")
        for i, fencer in enumerate(visible_fencers):
            gnuplot_script.write(u"'%s' using 1:%d with lines lw 5 title '%s'" %
                                 (data_file_name, fencer.columnindex, fencer.name))
            if i < len(visible_fencers) - 1:
                gnuplot_script.write(", ")
        gnuplot_script.close()
        try:
            call([parameters["path of gnuplot"], gnuplot_script_file_name])
        except OSError:
            raise ExternalProgramError(_(u'The program "gnuplot" wasn\'t found.  %s'
                                         u'However, it is needed for the plots.  '
                                         u'Please install it from http://www.gnuplot.info/.') %
                                       construct_supplement(parameters["path of gnuplot"]))
        os.remove(data_file_name)
        os.remove(gnuplot_script_file_name)
        if os.path.isfile(pdf_file_name):
            # Otherwise, convert doesn't generate a fresh PDF
            os.remove(pdf_file_name)
        try:
            call([parameters["path of convert"], postscript_file_name, "-rotate", "90", png_file_name])
            call([parameters["path of convert"], postscript_file_name, pdf_file_name])
        except OSError:
            raise ExternalProgramError(_(u'The program "convert" of ImageMagick wasn\'t found.  %s'
                                         u'However, it is needed for the plots.  '
                                         u'Please install it from http://www.imagemagick.org/.') %
                                       construct_supplement(parameters["path of convert"]))
        os.remove(postscript_file_name)
    if estimate_freshmen:
        return [fencer for fencer in fencers.values() if fencer.freshman]
    return visible_fencers


def expectation_value(first_fencer, second_fencer):
    """Returns the expected winning value of the first given fencer in a bout
    with the second.  The winning value is actually the fraction of won single
    points.  I.e., an expectation value of 0.7 means that the first fencer will
    make 70% of the points.

    :Parameters:
      - `first_fencer`: first fencer, for whom the expectation value is to be
        calculated
      - `second_fencer`: second fencer, opponent in the bout

    :type first_fencer: Fencer
    :type second_fencer: Fencer

    :Return:
      - expected winning value, with 0 S{<=} value S{<=} 1.

    :rtype: float
    """
    return 1 / (1 + 10 ** ((second_fencer.felo_rating_exact - first_fencer.felo_rating_exact) / 400.0))


def prognosticate_bout(first_fencer, second_fencer, fenced_to):
    """Estimates the most probable result of a bout.

    :Parameters:
      - `first_fencer`: first fencer, for whom the winning chance is to be
        calculated.
      - `second_fencer`: second fencer, opponent in the bout.
      - `fenced_to`: number of winning points in the bout.  Can be e.g. 5, 10, or
        15.

    :type first_fencer: Fencer
    :type second_fencer: Fencer
    :type fenced_to: int

    :Return:
      - Points of the first fencer, points of the second fencer, and the
        winning probability of the first fencer in percent.

    :rtype: int, int, int
    """
    expectation_first = expectation_value(first_fencer, second_fencer)
    if expectation_first > 0.5:
        points_first = fenced_to
        points_second = int(round(fenced_to * (1 / expectation_first - 1)))
    else:
        points_first = int(round(fenced_to / (1 / expectation_first - 1)))
        points_second = fenced_to
    for line in open(datapath + "/auf%d.dat" % fenced_to, 'r'):
        if line[0] != "#":
            result_first, winning_chance_first, __ = line.split()
            result_first, winning_chance_first = float(result_first), float(winning_chance_first)
            if abs(expectation_first - result_first) < 0.0051:
                break
    if points_first == points_second:
        if winning_chance_first > 0.5:
            points_second -= 1
        else:
            points_first -= 1
    return points_first, points_second, int(round(winning_chance_first * 100))


if __name__ == '__main__':
    """If called as a program, it interprets the command line parameters and
    calculates the resulting Felo rankings for each given Felo file.  It prints
    the result lists on the screen or writes them to a file.  It's a quick and
    simple way to calculate numbers, and it may be enough for many purposes.
    """
    import sys, optparse

    option_parser = optparse.OptionParser()
    option_parser.add_option("-p", "--plots", action="store_true", dest="plots",
                             help=_(u"Generate plots with the Felo ratings"), default=False)
    option_parser.add_option("-b", "--bootstrap", action="store_true", dest="bootstrap",
                             help=_(u"Try to estimate good initial values for all fencers"), default=False)
    option_parser.add_option("--max-cycles", type="int",
                             dest="max_cycles", help=_(u"Maximal iteration steps during bootstrapping."
                                                       "  Default: 1000"),
                             default=1000, metavar=_(u"NUMBER"))
    option_parser.add_option("--estimate-freshmen", action="store_true", dest="estimate_freshmen",
                             help=_(u"Try to estimate freshmen"), default=False)
    option_parser.add_option("--write-back", action="store_true", dest="write_back",
                             help=_(u"Write the new initial values back into the Felo file"),
                             default=False)
    option_parser.add_option("--version", action="store_true", dest="version",
                             help=_(u"Print out version number and copying information"),
                             default=False)
    option_parser.add_option("-o", "--output", type="string",
                             dest="output_file",
                             help=_(u"Name of the output file.  Default: output to the screen (stdout)"),
                             default=None, metavar=_(u"FILENAME"))
    options, felo_filenames = option_parser.parse_args()

    if options.version:
        print("Felo ratings " + distribution_version + _(u", revision %s") % __version__[11:-2])
        print()
        print((u"Copyright %s 2006 Torsten Bronger, Aachen, Germany") % u"©")
        print((u"""This is free software.  You may redistribute copies of it under the terms of
the MIT License <http://www.opensource.org/licenses/mit-license.php>.
There is NO WARRANTY, to the extent permitted by law."""))
        print()
        print((u"Written by Torsten Bronger <bronger@physik.rwth-aachen.de>."))
        sys.exit()
    try:
        if options.estimate_freshmen and options.bootstrap:
            raise Error(_(u"You cannot bootstrap and estimate freshmen at the same time."))
        if options.output_file:
            output_file = codecs.open(options.output_file, "w")
        else:
            output_file = sys.stdout
        for i, felo_filename in enumerate(felo_filenames):
            parameters, given_parameters, fencers, bouts = \
                parse_felo_file(codecs.open(felo_filename, encoding="utf-8"))
            resultslist = calculate_felo_ratings(parameters, fencers, bouts, options.plots,
                                                 options.estimate_freshmen,
                                                 options.bootstrap, options.max_cycles)
            if options.write_back and (options.bootstrap or options.estimate_freshmen):
                for fencer in fencers.values():
                    if (options.bootstrap and not fencer.freshman) or \
                            (options.estimate_freshmen and fencer.freshman):
                        fencer.initial_felo_rating = fencer.felo_rating
                write_back_fencers_to_file(felo_filename, fencers)
            if len(felo_filenames) > 1:
                if i >= 1: print(file=output_file)
                print(parameters["groupname"] + ":", file=output_file)
            for fencer in resultslist:
<<<<<<< HEAD
                print>>output_file, "    " + fencer.name + (19-len(fencer.name))*" " + "\t" + \
                    str(fencer.felo_rating)
    except Error, e:
        print>>sys.stderr, "felo_rating:", e.description
=======
                print("    " + fencer.name + (19 - len(fencer.name)) * " " + "\t" + str(fencer.felo_rating),
                      file=output_file)
    except Error as e:
        print("felo_rating:", e.description, file=sys.stderr)
>>>>>>> origin/Felo-Library-Integration
