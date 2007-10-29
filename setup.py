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

from distutils.core import setup
from ooopy.Version  import VERSION

description = []
f = open ('README')
f.next ()
f.next ()
for line in f :
    if line.strip () == 'Resources' :
        break
    description.append (line)

license = 'GNU Library or Lesser General Public License (LGPL)'

setup \
    ( name             = "OOoPy"
    , version          = VERSION
    , description      = "OOoPy: Modify OpenOffice.org documents in Python"
    , long_description = ''.join (description)
    , license          = license
    , author           = "Ralf Schlatterbeck"
    , author_email     = "rsc@runtux.com"
    , url              = "http://ooopy.sourceforge.net/"
    , packages         = ['ooopy']
    , platforms        = 'Any'
    , data_files       = [('share/ooopy' , ['test.sxw','run_doctest.py'])]
    , scripts          = ['ooo_as_text']
    , classifiers      = \
        [ 'Development Status :: 3 - Alpha'
        , 'License :: OSI Approved :: ' + license
        , 'Operating System :: OS Independent'
        , 'Programming Language :: Python'
        , 'Topic :: Internet :: WWW/HTTP :: Dynamic Content'
        , 'Topic :: Office/Business :: Office Suites'
        , 'Topic :: Text Processing'
        , 'Topic :: Text Processing :: General'
        , 'Topic :: Text Processing :: Markup :: XML'
        , 'Topic :: Text Editors :: Word Processors'
        ]
    )
