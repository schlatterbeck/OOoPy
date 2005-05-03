#!/usr/bin/env python
# -*- coding: iso-8859-1 -*-
# Copyright (C) 2005 Dr. Ralf Schlatterbeck Open Source Consulting.
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

import sys
from ooopy.OOoPy     import OOoPy
from ooopy.Transform import namespaces

names = {}

def cleantag (tag) :
    ns, t = tag.split ('}')
    return ':'.join ((names [ns [1:]], t))

def pretty (n, indent = 0) :
    s = ["    " * indent]
    s.append (cleantag (n.tag))
    attrkeys = n.attrib.keys ()
    attrkeys.sort ()
    for a in attrkeys :
        s.append (' %s="%s"' % (cleantag (a), n.attrib [a]))
    print ''.join (s).encode ('utf-8')
    for sub in n :
        pretty (sub, indent + 1)

if __name__ == '__main__' :
    for n in namespaces :
        names [namespaces [n]] = n
    ooofile = 'content.xml'
    idx = 1
    if len (sys.argv) >= 2 :
        if sys.argv [1].startswith ('--file') :
            ooofile = sys.argv [1].split ('=') [1]
            idx = 2
    for f in sys.argv [idx:] :
        o = OOoPy (infile = f)
        e = o.read (ooofile)
        pretty (e.getroot ())
        o.close ()
