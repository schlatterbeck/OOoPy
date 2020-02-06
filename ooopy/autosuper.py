#!/usr/bin/env python
# Copyright (C) 2020 Dr. Ralf Schlatterbeck Open Source Consulting.
# Reichergasse 131, A-3411 Weidling.
# Web: http://www.runtux.com Email: office@runtux.com
# All rights reserved
# ****************************************************************************
#
# This library is free software; you can redistribute it and/or modify
# it under the terms of the GNU Library General Public License as
# published by the Free Software Foundation; either version 2 of the
# License, or (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Library General Public License for more details.
#
# You should have received a copy of the GNU Library General Public
# License along with this program; if not, write to the Free Software
# Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.
# ****************************************************************************

def with_metaclass (meta, *bases) :
    class metaclass (meta) :
        __call__ = type.__call__
        __init__ = type.__init__
        def __new__ (cls, name, this_bases, d) :
            if this_bases is None :
                return type.__new__ (cls, name, (), d)
            return meta (name, bases, d)
    return metaclass ('temporary_class', None, {})
# end def with_metaclass

class _autosuper (type) :
    def __init__ (cls, name, bases, dict) :
        super   (_autosuper, cls).__init__ (name, bases, dict)
        setattr (cls, "_%s__super" % name, super (cls))
    # end def __init__
# end class _autosuper

class autosuper (with_metaclass (_autosuper)) :
    """ Autosuper magic
    >>> from OOoPy import autosuper
    >>> class X (autosuper, dict) :
    ...     def __init__ (self, *args, **kw) :
    ...         return self.__super.__init__ (*args, **kw)
    ...
    >>> X((x,1) for x in range(3))
    {0: 1, 1: 1, 2: 1}
    >>> class Y (autosuper) :
    ...     def __repr__ (self) :
    ...         return "class Y"
    ...
    >>> Y((x,1) for x in range(23))
    class Y
    """
    __metaclass__ = _autosuper
    def __init__ (self, *args, **kw) :
        try :
            oc =  self.__super.__init__.__objclass__
        except AttributeError :
            oc = None
        if oc is object :
            self.__super.__init__ ()
        else :
            self.__super.__init__ (*args, **kw)
    # end def __init__
# end class autosuper
