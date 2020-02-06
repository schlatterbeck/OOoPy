OOoPy: Modify OpenOffice.org documents in Python
================================================

:Author: Ralf Schlatterbeck <rsc@runtux.com>

OpenOffice.org (OOo) documents are ZIP archives containing several XML
files.  Therefore it is easy to inspect, create, or modify OOo
documents. OOoPy is a library in Python for these tasks with OOo
documents. To not reinvent the wheel, OOoPy uses an existing XML
library, ElementTree by Fredrik Lundh (which is in the Python standard
library for quite some time now). OOoPy is a thin wrapper around
ElementTree using Python's ZipFile to read and write OOo documents.

In addition to being a wrapper for ElementTree, OOoPy contains a
framework for applying XML transforms to OOo documents. Several
Transforms for OOo documents exist, e.g., for changing OOo fields (OOo
Insert-Fields menu) or using OOo fields for a mail merge application.
Some other transformations for modifying OOo settings and meta
information are also given as examples.

Applications like this come in handy in applications where calling
native OOo is not an option, e.g., in server-side Web applications.

If the mailmerge transform doesn't work for your document: The OOo
format is well documented but there are ordering constraints in the body
of an OOo document.
I've not yet figured out all the tags and their order in the
OOo body. Individual elements in an OOo document (like e.g., frames,
sections, tables) need to have their own unique names.  After a mailmerge,
there are duplicate names for some items. So far I'm renumbering only
frames, sections, and tables. See the renumber objects at the end of
ooopy/Transforms.py. So if you encounter missing parts of the mailmerged
document, check if there are some renumberings missing or send me a `bug
report`_.

.. _`bug report`: http://ooopy.sourceforge.net/#reporting-bugs

There is currently not much documentation except for a python doctest in
OOoPy.py and Transformer.py and the command-line utilities_.
For running these test, after installing
ooopy (assuming here you installed using python into /usr/local)::

 cd /usr/local/share/ooopy
 python run_doctest.py /usr/local/lib/python2.X/site-packages/ooopy/Transformer.py
 python run_doctest.py /usr/local/lib/python2.X/site-packages/ooopy/OOoPy.py

Both should report no failed tests.

Usage
-----

There were some slight changes to the API when supporting the open
document format introduced with OOo 2.0. See below if you get a traceback
when upgrading from an old version.

See the online documentation, e.g.::

 % python
 >>> from ooopy.OOoPy import OOoPy
 >>> help (OOoPy)
 >>> from ooopy.Transformer import Transformer
 >>> help (Transformer)

Help, I'm getting an AssertionError traceback from Transformer, e.g.::

 Traceback (most recent call last):
   File "./replace.py", line 17, in ?
     t = Transformer(Field_Replace(replace = replace_dictionary))
   File "/usr/local/lib/python2.4/site-packages/ooopy/Transformer.py", line 1226, in __init__
     assert (mimetype in mimetypes)
 AssertionError

The API changed slightly when implementing handling of different
versions of OOo files. Now the first parameter you pass to the
Transformer constructor is the mimetype of the OpenOffice.org document
you intend to transform. The mimetype can be fetched from another opened
OOo document, e.g.::

  ooo = OOoPy (infile = 'test.odt', outfile = 'test_out.odt')
  t = Transformer(ooo.mimetype, ...

Usage of Command-Line Utilities
-------------------------------

A, well, there are command-line _`utilities` now:

- ooo_cat for concatenating several OOo files into one
- ooo_grep to do equivalent of grep -l on OOo files -- only runs on
  Unix-like operating systems, probably only with the GNU version of grep
  (it's a shell-script using ooo_as_text) Note that the -l option of
  grep only prints the matching filenames.
- ooo_fieldreplace for replacing fields in an OOo document
- ooo_mailmerge for doing a mailmerge from a template OOo document and a
  CSV (comma separated values) input
- ooo_as_text for getting the text from an OOo-File (e.g., for doing a
  "grep" on the output).
- ooo_prettyxml for pretty-printing the XML nodes of one of the XML
  files inside an OOo document. Mainly useful for debugging.

All utilities take a ``--help`` option.

Resources
---------

Project information and download from `Sourceforge main page`_

.. _`Sourceforge main page`: http://sourceforge.net/projects/ooopy/

You need at least version 2.7 of python. Now also tested with 3.5, will
probably work with later versions, too.

For documentation about the OOo XML file format, see the book by
J. David Eisenberg called `OASIS OpenDocument Essentials`_ which is
under the Gnu Free Documentation License and is also available `in
print`_.  For a reference document you may want to check out the `XML
File Format Specification`_ (PDF) by OpenOffice.org.

A german page for OOoPy exists at `runtux.com`_

.. _`ElementTree Library`: http://effbot.org/downloads/#elementtree
.. _`OASIS OpenDocument Essentials`: http://books.evc-cit.info/
.. _`in print`:
   http://www.lulu.com/product/paperback/oasis-opendocument-essentials/392512
.. _`XML File Format Specification`:
   http://xml.openoffice.org/xml_specification.pdf
.. _`runtux.com`: http://www.runtux.com/ooopy.html

Reporting Bugs
--------------
Please use the `Sourceforge Bug Tracker`_ and

 - attach the OOo document that reproduces your problem
 - give a short description of what you think is the correct behaviour
 - give a description of the observed behaviour
 - tell me exactly what you did.

.. _`Sourceforge Bug Tracker`:
    http://sourceforge.net/tracker/?group_id=134329&atid=729727

Changes
-------

Version 2.0: Port to Python3

Still working with lower python versions but I'm only able to test with
2.7, nothing earlier. Bug fixes where sometimes multiple Set_Attribute
transformations (e.g. when concatenating OO documents) would be applied
to the same tag/attribute combination. Also fix a bug with the
computation of default tabulators when concatenating documents.

Version 1.11: Small Bug fix ooo_mailmerge

Now ooo_mailmerge uses the delimiter option, it was ignored before.
Thanks to Bob Danek for report and test.

 - Fix setting csv delimiter in ooo_mailmerge

Version 1.10: Fix table styles when concatenating

Now ooo_cat fixes tables styles when concatenating (renaming): We
optimize style usage by re-using existing styles. But for some table
styles the original names were not renamed to the re-used ones.
Fixes SF Bug 10, thanks to Claudio Girlanda for reporting.

 - Fix style renaming for table styles when concatenating documents
 - Add some missing namespaces (ooo 2009)

Version 1.9: Add Picture Handling for Concatenation

Now ooo_cat supports pictures, thanks to Antonio Sánchez for reporting
that this wasn't working.

 - Add a list of filenames + contents to Transformer
 - Update this file-list in Concatenate
 - Add Manifest_Append transform to update META-INF/manifest.xml with
   added filenames
 - Add hook in OOoPy for adding files
 - Update tests
 - Update ooo_cat to use new transform
 - This is the first release after migration of the version control from
   Subversion to GIT

Version 1.8: Minor bugfixes

Distribute a missing file that is used in the doctest. Fix directory
structure. Thanks to Michael Nagel for suggesting the change and
reporting the bug.

 - The file ``testenum.odt`` was missing from MANIFEST.in
 - All OOo files and other files needed for testing are now in the
   subdirectory ``testfiles``.
 - All command line utilities are now in subdirectory ``bin``.

Version 1.7: Minor feature additions

Add --newlines option to ooo_as_text: With this option the paragraphs in
the office document are preserved in the text output.
Fix assertion error with python2.7, thanks to Hans-Peter Jansen for the
report. Several other small fixes for python2.7 vs. 2.6.

 - add --newlines option to ooo_as_text
 - fix assertion error with python2.7 reported by Hans-Peter Jansen
 - fix several deprecation warnings with python2.7
 - remove zip compression sizes from regression test: the compressor in
   python2.7 is better than the one in python2.6

Version 1.6: Minor bugfixes

Fix compression: when writing new XML-files these would be stored
instead of compressed in the OOo zip-file resulting in big documents.
Thanks to Hans-Peter Jansen for the patch. Add copyright notice to
command-line utils (SF Bug 2650042). Fix mailmerge for OOo 3.X lists (SF
Bug 2949643).

 - fix compression flag, patch by Hans-Peter Jansen
 - add regression test to check for compression
 - now release ooo_prettyxml -- I've used this for testing for quite
   some time, may be useful to others
 - Add copyright (LGPL) notice to command-line utilities, fixes SF Bug
   2650042
 - OOo 3.X adds xml:id tags to lists, we now renumber these in the
   mailmerge app., fixes SF Bug 2949643

Version 1.5: Minor feature enhancements

Add ooo_grep to search for OOo files containing a pattern. Thanks to
Mathieu Chauvinc for the reporting the problems with modified
manifest.xml.
Support python2.6, thanks to Erik Myllymaki for reporting and anonymous
contributor(s) for confirming the bug.

 - New shell-script ooo_grep (does equivalent to grep -l on OOo Files)
 - On deletion of an OOoPy object close it explicitly (uses __del__)
 - Ensure mimetype is the first element in the resulting archive, seems
   OOo is picky about this.
 - When modifying the manifest the resulting .odt file could not be
   opened by OOo. So when modifying manifest make sure the manifest
   namespace is named "manifest" not something auto-generated by
   ElementTree. I consider this a bug in OOo to require this. This now
   uses the _namespace_map of ElementTree and uses the same names as OOo
   for all namespaces. The META-INF/manifest.xml is now in the list of
   files to which Transforms can be applied.
 - When modifying (or creating) archive members, we create the OOo
   archive as if it was a DOS system (type fat) and ensure we use the
   current date/time (UTC). This also fixes problems with file
   permissions on newer versions of pythons ZipFile.
 - Fix for python2.6 behavior that __init__ of object may not take any
   arguments. Fixes SF Bug 2948617.
 - Finally -- since OOoPy is in production in some projects -- change the
   development status to "Production/Stable".

Version 1.4: Minor bugfixes

Fix Doctest to hopefully run on windows. Thanks to Dani Budinova for
testing thoroughly under windows.

 - Open output-files in "wb" mode instead of "w" in doctest to not
   create corrupt OOo documents on windows.
 - Use double quotes for arguments when calling system, single quotes
   don't seem to work on windows.
 - Dont use redirection when calling system, use -i option for input
   file instead. Redirection seems to be a problem on windows.
 - Explicitly call the python-interpreter, running a script directly is
   not supported on windows.

Version 1.3: Minor bugfixes

Regression-test failed because some files were not distributed.
Fixes SF Bugs 1970389 and 1972900.

 - Fix MANIFEST.in to include all files needed for regression test
   (doctest).

Version 1.2: Major feature enhancements

Add ooo_fieldreplace, ooo_cat, ooo_mailmerge command-line utilities. Fix
ooo_as_text to allow specification of output-file. Note that handling of
non-seekable input/output (pipes) for command-line utils will work only
starting with python2.5. Minor bug-fix when concatenating documents. 

 - Fix _divide (used for dividing body into parts that must keep
   sequence). If one of the sections was empty, body parts would change
   sequence.
 - Fix handling of cases where we don't have a paragraph (only list) elements
 - Implement ooo_cat
 - Fix ooo_as_text to include more command-line handling
 - Fix reading/writing stdin/stdout for command-line utilities, this
   will work reliably (reading/writing non-seekable input/output like,
   e.g., pipes) only with python2.5
 - implement ooo_fieldreplace and ooo_mailmerge

Version 1.1: Minor bugfixes

Small Documentation changes

 - Fix css stylesheet
 - Link to SF logo for Homepage
 - Link to other information updated
 - Version numbers in documentation fixed
 - Add some checks for new API -- first parameter of Transformer is checked now
 - Ship files needed for running the doctest and explain how to run it
 - Usage section

Version 1.0: Major feature enhancements

Now works with version 2.X of OpenOffice.org. Minor API changes.

 - Tested with python 2.3, 2.4, 2.5
 - OOoPy now works for OOo version 1.X and version 2.X
 - New attribute mimetype of OOoPy -- this is automatically set when
   reading a document, and should be set when writing one.
 - renumber_all, get_meta, set_meta are now factory functions that take
   the mimetype of the open office document as a parameter.
 - Since renumber_all is now a function it will (correctly) restart
   numbering for each new Attribute_Access instance it returns.
 - Built-in elementtree support from python2.5 is used if available
 - Fix bug in optimisation of original document for concatenation
