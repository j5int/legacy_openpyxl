from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
# @license: http://www.opensource.org/licenses/mit-license.php
# @author: see AUTHORS file
from numbers import Number

from legacy_openpyxl.shared.units import short_color

from .reference import Reference


class Series(object):
    """ a serie of data and possibly associated labels """

    MARKER_NONE = 'none'
    _title = None
    _legend = None

    def __init__(self, values, title=None, labels=None, color=None,
                 xvalues=None, legend=None):

        self.marker = Series.MARKER_NONE
        self.values = values
        self.xvalues = xvalues
        self.labels = labels
        self.title = title
        self.error_bar = None
        if legend is not None:
            self.legend = legend

    @property
    def title(self):
        if self._title is not None:
            return self._title
        if self.legend is not None:
            return self.legend.values[0]

    @title.setter
    def title(self, value):
        self._title = value

    @property
    def legend(self):
        return self._legend

    @legend.setter
    def legend(self, value):
        from warnings import warn
        warn("Series titles can be set directly using series.title. Series legend will be removed in 2.0")
        value.data_type = 's'
        self._legend = value

    @property
    def color(self):
        return getattr(self, "_color", None)

    @color.setter
    def color(self, color):
        if color is None:
            raise ValueError("Colors must be strings of the format XXXXX")
        self._color = short_color(color)

    @property
    def values(self):
        """Return values from underlying reference"""
        return self._values

    @values.setter
    def values(self, reference):
        """Assign values from reference to serie"""
        if reference is not None:
            if not isinstance(reference, Reference):
                raise TypeError("Series values must be a Reference instance")
            self._values = reference.values
        else:
            self._values = None
        self.reference = reference

    @property
    def xvalues(self):
        """Return xvalues"""
        return self._xvalues

    @xvalues.setter
    def xvalues(self, reference):
        if reference is not None:
            if not isinstance(reference, Reference):
                raise TypeError("Series xvalues must be a Reference instance")
            self._xvalues = reference.values
        else:
            self._xvalues = None
        self.xreference = reference

    @property
    def labels(self):
        """Return values from reference set as label"""
        return self._labels

    @labels.setter
    def labels(self, reference):
        if reference is not None:
            if not isinstance(reference, Reference):
                raise TypeError("Series labels must be a Reference instance")
            reference.values
            self._labels = reference
        else:
            self._labels = None

    def max(self, attr='values'):
        """
        Return the maximum value for numeric series.
        NB None has a value of u'' which is ignored
        """
        values = getattr(self, attr)
        if self.error_bar:
            values = self._error_bar_values
        cleaned = [v for v in values if isinstance(v, Number)]
        if cleaned:
            return max(cleaned)

    def min(self, attr='values'):
        """
        Return the minimum value for numeric series
        NB None has a value of u'' which is ignored
        """
        values = getattr(self, attr)
        if self.error_bar:
            values = self._error_bar_values
        cleaned = [v for v in values if isinstance(v, Number)]
        if cleaned:
            return min(cleaned)

    @property
    def _error_bar_values(self):
        """Documentation required here"""
        err_cache = self.error_bar.values
        vals = [v + err_cache[i] \
            for i, v in enumerate(self.values)]
        return vals

    def get_min_max(self):
        """Legacy method. Replaced by properties"""
        return self.min(), self.max()

    def __len__(self):

        return len(self.values)

# backwards compatibility
Serie = Series
