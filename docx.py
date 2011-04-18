#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''Pragmatic attempt to write OOXML (MS Word 2007) files from python.

The module is still very incomplete as new features are implemented only when
I need them :-)

.. codeauthor:: Ole Streicher <ole@aip.de>
'''

import sys
import os
import shutil
from xml.dom import minidom
import tempfile
import zipfile
import codecs

ns = { 
    'w'  :'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wx' :'http://schemas.openxmlformats.org/wordprocessingml/2006/auxHint',
    'wp' :'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a'  :'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic':'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'r'  :'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'vt' :'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
    }
       

try:
    import matplotlib.figure
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    _have_matplotlib = True
except:
    _have_matplotlib = False

def open(fname = None, mode = 'copyonwrite'):
    '''Create a new (word) document, or load an existing document.
    
    fname: File name
    mode: Open mode. 'copyonwrite' (default), 'update', 'append' or 'create'.
    '''
    return Document(fname, mode)

class CustomProperty(object):
    '''Field of user settable properties.

    These properties may be used to fill defined fields of a document with
    custom values. Examples are the author, the document title, or a document
    id.
    '''
    def __init__(self, parent, pfile = None):
        self.parent = parent
        if pfile is not None:
            self.doc = minidom.parse(pfile)
        else:
            self.doc = minidom.Document()
            p = self.doc.createElement('Properties')
            p.setAttribute('xmlns', 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties')
            p.setAttribute('xmlns:vt', ns['vt'])
            self.doc.appendChild(p)
            
    def _get_TextNode(self, key, create = False):
        max_pid = 1
        for n in self.doc.getElementsByTagName('property'):
            max_pid = max(max_pid, int(n.getAttribute('pid')))
            if n.getAttribute('name') == key:
                break
        else:
            if create:
                n = self.doc.createElement('property')
                n.setAttribute('name', key)
                n.setAttribute('pid', '%i' % (max_pid+1))
                self.doc.documentElement.appendChild(n)
            else:
                return None
        vl = n.getElementsByTagName('vt:lpwstr')
        if len(vl) > 0:
            v = vl[0]
        else:
            v = self.doc.createElement('vt:lpwstr')
            n.appendChild(v)
        for t in v.childNodes:
            if t.nodeType == v.TEXT_NODE:
                break
        else:
            t = self.doc.createTextNode('')
            v.appendChild(t)
        return t

    def _update_field(self, element, key, value):
        instr = ' DOCPROPERTY  "%s"  \\* MERGEFORMAT ' % key
        # Replace "Simple" properties
        for n in element.getElementsByTagName('w:fldSimple'):
            if n.getAttribute('w:instr').strip() != instr.strip():
                continue
            wp = n.parentNode
            nwr = wp.ownerDocument.createElement('w:r')
            flds = nwr.ownerDocument.createElement('w:fldSimple')
            flds.setAttribute('w:instr', instr)
            wr = n.getElementsByTagName('w:r')
            if len(wr) > 0:
                for rpr in wr[0].getElementsByTagName('w:rPr'):
                    wr[0].removeChild(rpr)
                    flds.appendChild(rpr)
            wt = flds.ownerDocument.createElement('w:t')
            wt.appendChild(wp.ownerDocument.createTextNode(value))
            nwr.appendChild(wt)
            flds.appendChild(nwr)
            wp.replaceChild(flds, n)

        # Replace "Complex" properties.
        for n in element.getElementsByTagName('w:instrText'):
            if n.firstChild is None:
                continue
            if n.firstChild.data.strip() != instr.strip():
                continue
            wr = n.parentNode
            wp = wr.parentNode
            wrs = []
            wr0 = wr.previousSibling
            while wr0 is not None:
                wrs.append(wr0)
                if wr0.nodeType == wr0.ELEMENT_NODE:
                    flds = wr0.getElementsByTagName('w:fldChar')
                    if len(flds) > 0:
                        break
                wr0 = wr0.previousSibling
            wr0 = wr.nextSibling
            while wr0 is not None:
                wrs.append(wr0)
                if wr0.nodeType == wr0.ELEMENT_NODE:
                    flds = wr0.getElementsByTagName('w:fldChar')
                    if len(flds) > 0:
                        if flds[0].getAttribute('w:fldCharType') == 'end':
                            break
                wr0 = wr0.nextSibling
            for wr0 in wrs:
                wp.removeChild(wr0)
                wr0.unlink()
            nwr = wp.ownerDocument.createElement('w:r')
            flds = nwr.ownerDocument.createElement('w:fldSimple')
            flds.setAttribute('w:instr', instr)
            for rpr in wr.getElementsByTagName('w:rPr'):
                wr.removeChild(rpr)
                flds.appendChild(rpr)
            
            wt = flds.ownerDocument.createElement('w:t')
            wt.appendChild(wp.ownerDocument.createTextNode(value))
            nwr.appendChild(wt)
            flds.appendChild(nwr)
            wp.replaceChild(flds, wr)
                    
    def __getitem__(self, key):
        v = self._get_TextNode(key)
        return v.data or None

    def __setitem__(self, key, value):
        self._get_TextNode(key, True).data = value
        self._update_field(self.parent.body, key, value)
        self._update_field(self.parent.header.documentElement, key, value)

    def __iter__(self):
        return iter(n.getAttribute('name') 
                    for n in self.doc.getElementsByTagName('property'))

class Settings(object):
    '''XML tree containing all document settings.
    '''
    def __init__(self, sfile = None):
        if sfile is not None:
            self.doc = minidom.parse(sfile)
        else:
            self.doc = minicom.Document()
            p = self.doc.createElement('w:settings')
            p.setAttribute('xmlns:w', ns['w'])
            self.doc.appendChild(p)

class Document(object):
    '''Main document.

    This contains the main document as well as all necessary subdocuments.
    '''
    def __init__(self, fname = None, mode = 'copyonwrite'):
        '''Create a new (word) document, or load an existing document.

        fname: File name
        mode: Open mode. 'copyonwrite' (default), 'update', 'append' or 
              'create'.
        '''
        self.fname = fname
        self.mode = mode
        self.tmpdir = tempfile.mkdtemp(prefix = 'word')
        worddir = os.path.join(self.tmpdir, 'word')
        self.mediadir = os.path.join(worddir, 'media')
        os.mkdir(worddir)
        os.mkdir(self.mediadir)
        os.mkdir(os.path.join(self.tmpdir, 'docProps'))
        os.mkdir(os.path.join(self.tmpdir, '_rels'))
        os.mkdir(os.path.join(worddir, '_rels'))
        self.media = { }
        self.styles = { 'heading 1':'berschrift1', 'heading 2':'berschrift2', 
                        'heading 3':'berschrift3', 'heading 4':'berschrift4',
                        'heading 5':'berschrift5', 'heading 6':'berschrift6',
                        'heading 7':'berschrift7', 'heading 8':'berschrift8',
                        'caption':'Beschriftung', 'Normal':'Standard' }
        if fname is None or mode is 'create':
            self._createdefault()
        elif os.path.exists(fname):
            self._load(fname)
        elif mode is 'create' or mode is 'append':
            self._createdefault()
        else:
            raise IOError('%s does not exist' % fname)

    def _createdefault(self):
        doc = minidom.Document()
        wdoc = doc.createElement('w:document')
        wdoc.setAttribute('xmlns:w', ns['w'])
        wdoc.setAttribute('xmlns:wx', ns['wx'])
        wdoc.setAttribute('xmlns:wp', ns['wp'])
        wdoc.setAttribute('xmlns:a', ns['a'])
        wdoc.setAttribute('xmlns:r', ns['r'])
        doc.appendChild(wdoc)
        self.body = doc.createElement('w:body')
        wdoc.appendChild(self.body)
        relations = minidom.Document()
        relsElem = relations.createElement("Relationships")
        relsElem.setAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships")
        relations.appendChild(relsElem)
        relshipElem = relations.createElement("Relationship")
        relshipElem.setAttribute("Id", "myrId1")
        relshipElem.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
        relshipElem.setAttribute("Target", "word/document.xml")
        relsElem.appendChild(relshipElem)
        self._xmlwrite(relations, os.path.join('_rels', '.rels'))
        content = minidom.Document()
        cTypes = content.createElement('Types')
        cTypes.setAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types')
        mime_types = { 'png':'image/png', 
                       'jpg':'image/jpeg', 
                       'xml':'application/xml', 
                       'rels':'application/vnd.openxmlformats-package.relationships+xml' }
        for e, c in mime_types.items():
            type = content.createElement('Default')
            type.setAttribute('Extension', e)
            type.setAttribute('ContentType', c)
            cTypes.appendChild(type)
        overrides = { '/word/document.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml' }
        for p, c in overrides.items():
            type = content.createElement('Override')
            type.setAttribute('PartName', p)
            type.setAttribute('ContentType', c)
            cTypes.appendChild(type)
        content.appendChild(cTypes)
        self._xmlwrite(content, '[Content_Types].xml')
        self.property = None
        self.settings = None
        self.header = None
        self.numberings = None

    def _load(self, filename):
        zip = zipfile.ZipFile(filename, 'r')
        zip.extractall(self.tmpdir)
        wdoc = minidom.parse(zip.open('word/document.xml')).documentElement
        wdoc.setAttribute('xmlns:wx', ns['wx'])
        wdoc.setAttribute('xmlns:a', ns['a'])
        self.body = wdoc.getElementsByTagName('w:body')[0]
        relfile = zip.open('word/_rels/document.xml.rels')
        for n in minidom.parse(relfile).getElementsByTagName('Relationship'):
            self.media[n.getAttribute('Id')] = ( n.getAttribute('Target'), 
                                                 n.getAttribute('Type') )
        sdoc = minidom.parse(zip.open('word/styles.xml'))
        for s in sdoc.getElementsByTagName('w:style'):
            style_id = s.getAttribute('w:styleId')
            n = s.getElementsByTagName('w:name')[0]
            style_name = n.getAttribute('w:val')
            self.styles[style_name] = style_id
        try:
            self.property = CustomProperty(self, 
                                           zip.open('docProps/custom.xml'))
        except:
            self.property = None
        try:
            self.settings = Settings(zip.open('word/settings.xml'))
        except:
            self.settings = None
        try:
            self.header = minidom.parse(zip.open('word/header1.xml'))
        except:
            self.header = None
        try:
            self.numberings = Numbering(zip.open('word/numbering.xml'))
        except:
            self.numberings = None
        zip.close()

    def __iadd__(self, other):
        if isinstance(other, (str, unicode)):
            self += Paragraph(other)
        elif isinstance(other, list):
            self += Table(other)
        elif _have_matplotlib and isinstance(other, matplotlib.figure.Figure):
            self += MatplotlibFigure(other)
        else:
            other.append_to(self, self.body)
        return self

    def appendMedia(self, fname):
        '''Append an external file.

        The file will be copied into the document tree and a reference to the
        file is returned.
        '''
        id = 'rId%i' % (len(self.media) + 1)
        type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
        self.media[id] = ('media/%s' % fname, type)
        return id

    def numbering(self, level, indent, hanging, bullet):
        if self.numberings is None:
            self.numberings = Numbering()
        return self.numberings.add(level, indent, hanging, bullet)

    def _xmlwrite(self, document, path):
        fp = codecs.open(os.path.join(self.tmpdir, path), 'w', 'utf-8')
        document.writexml(fp, encoding = "UTF-8")
#        document.writexml(fp, "", "  ", "\n", encoding = "UTF-8")
        fp.close()

    def writeto(self, filename):
        '''Write the document to a file. The file will be overwritten without
        warning.
        '''
        self._xmlwrite(self.body.ownerDocument,
                       os.path.join('word', 'document.xml'))
        if self.property is not None:
            self._xmlwrite(self.property.doc, 
                           os.path.join('docProps', 'custom.xml'))
        if self.settings is not None:
            self._xmlwrite(self.settings.doc, 
                           os.path.join('word', 'custom.xml'))
        if self.header is not None:
            self._xmlwrite(self.header, 
                           os.path.join('word', 'header1.xml'))
        if self.numberings is not None:
            self._xmlwrite(self.numberings.doc, 
                           os.path.join('word', 'numbering.xml'))
        relations = minidom.Document()
        relsElem = relations.createElement("Relationships")
        relsElem.setAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships")
        relations.appendChild(relsElem)
        for id, (target, type) in self.media.items():
            relshipElem = relations.createElement("Relationship")
            relshipElem.setAttribute("Id", id)
            relshipElem.setAttribute("Type", type)
            relshipElem.setAttribute("Target", target)
            relsElem.appendChild(relshipElem)
        self._xmlwrite(relations, 
                       os.path.join('word', "_rels", 'document.xml.rels'))
        f = zipfile.ZipFile(filename, 'w', zipfile.ZIP_DEFLATED)
        for dirpath,dirnames,filenames in os.walk(self.tmpdir):
            for filename in filenames:
                f.write(os.path.join(dirpath,filename),
                        os.path.join(dirpath,filename).replace(self.tmpdir,''))
        f.close()

    def flush(self):
        '''Flush all changes to disk.
        
        This works only if the document was opened in 'update' or 'append'
        mode.
        '''
        if ((self.mode is 'update' or self.mode is 'append') 
            and self.fname is not None):
            self.writeto(self.fname)

    def close(self):
        '''Close the document and clean up the file space.
        
        In 'update'/'append' mode, the changes are flushed to the document
        file.
        '''
        self.flush()
        shutil.rmtree(self.tmpdir)

    def __del__(self):
        shutil.rmtree(self.tmpdir)

class Text(object):
    '''Structure containing some text with the same formatting options.
    
    '''
    def __init__(self, content, bold = None, italic = None, underline = None):
        self.content = content if isinstance(content, unicode) \
            else unicode(content)
        self.bold = bold
        self.italic = italic
        self.underline = underline

    def append_to(self, doc, target):
        r = target.ownerDocument.createElement('w:r')
        style = False
        rpr = target.ownerDocument.createElement('w:rPr')
        if self.bold is not None:
            style = True
            wb = target.ownerDocument.createElement('w:b')
            if not self.bold:
                wb.setAttribute('w:val', 'off')
            rpr.appendChild(wb)
        if self.italic is not None:
            style = True
            wi = target.ownerDocument.createElement('w:i')
            if not self.italic:
                wi.setAttribute('w:val', 'off')
            rpr.appendChild(wi)
        if self.underline is not None:
            style = True
            wu = target.ownerDocument.createElement('w:u')
            styles = {
                0:'off', False:'none',
                1:'single', True:'single', '_':'single',
                2:'double', '=':'double',
                '#':'thick',
                '-':'words',
                ',':'dash',
                '.':'dotted',
                ';':'dot-dash',
                }
            wu.setAttribute('w:val', styles.get(self.underline, 'none'))
            rpr.appendChild(wu)
        if style:
            r.appendChild(rpr)
        if self.content.startswith(' '):
            t = target.ownerDocument.createElement('w:t')
            t.setAttribute('xml:space', 'preserve')
            t.appendChild(target.ownerDocument.createTextNode(' '))
            r.appendChild(t)
        t = target.ownerDocument.createElement('w:t')
        t.appendChild(target.ownerDocument.createTextNode(self.content.strip()))
        r.appendChild(t)
        if self.content.endswith(' '):
            t = target.ownerDocument.createElement('w:t')
            t.setAttribute('xml:space', 'preserve')
            t.appendChild(target.ownerDocument.createTextNode(' '))
            r.appendChild(t)
        target.appendChild(r)

class Paragraph(object):
    '''Structure containing one text paragraph.
    '''
    def __init__(self, content = None, style = None, align = None):
        self.style = style
        self.align = align
        self.content = [ ]
        if content:
            self.__iadd__(content)

    def __iadd__(self, other):
        self.content.append(other if isinstance(other, Text) else Text(other))

    def append_to(self, doc, target, indent = None, numbering = None):
        p = target.ownerDocument.createElement('w:p')
        if self.style is not None or self.align is not None or numbering is not None or indent is not None:
            pPr = target.ownerDocument.createElement('w:pPr')
            if self.style:
                pStyle = target.ownerDocument.createElement('w:pStyle')
                pStyle.setAttribute('w:val', doc.styles[self.style])
                pPr.appendChild(pStyle)
            alignments = {
                'l':'left', 'left':'left', '<':'left',
                'r':'right', 'right':'right', '>':'right',
                'c':'center', 'center':'center', 
                'b':'both', 'block':'both', 'both':'both', '=':'both',
                }
            a = alignments.get(self.align)
            if a:
                pjc = target.ownerDocument.createElement('w:jc')
                pjc.setAttribute('w:val', a)
                pPr.appendChild(pjc)
            if numbering:
                numpr = target.ownerDocument.createElement('w:numPr')
                ilvl = target.ownerDocument.createElement('w:ilvl')
                ilvl.setAttribute('w:val', '%i' % numbering[0])
                numpr.appendChild(ilvl)
                numid = target.ownerDocument.createElement('w:numId')
                numid.setAttribute('w:val', '%i' % numbering[1])
                numpr.appendChild(numid)
                pPr.appendChild(numpr)
            if indent:
                ind = target.ownerDocument.createElement('w:ind')
                ind.setAttribute('w:left', '%i' % indent)
                pPr.appendChild(ind)
            p.appendChild(pPr)
        for c in self.content:
            c.append_to(doc, p)
        target.appendChild(p)

class Header(Paragraph):
    '''Caption header.
    '''
    def __init__(self, level, content):
        Paragraph.__init__(self, content, style = 'heading %i' % level)

class Counter(Text):
    '''Counter, for table and figure captions
    '''
    def __init__(self, name, bold = None, italic = None, underline = None):
        Text.__init__(self, '0', bold, italic, underline)
        self.name = name

    def get_counter(self, target):
        instr = ' SEQ %s \* ARABIC ' % self.name
        return len([n for n in target.getElementsByTagName('w:fldSimple')
                    if n.getAttribute('w:instr').strip() == instr.strip()]) + 1

    def append_to(self, doc, target):
        num = self.get_counter(target.ownerDocument)
        fld = target.ownerDocument.createElement('w:fldSimple')
        fld.setAttribute('w:instr', ' SEQ %s \* ARABIC ' % self.name)
        self.content = '%i' % num
        Text.append_to(self, doc, fld)
        target.appendChild(fld)

class Caption(Paragraph):
    '''Table or figure caption
    '''
    def __init__(self, name, content, style = 'caption'):
        Paragraph.__init__(self, content, style)
        self.name = name
        self.counter = Counter(name)

    def append_to(self, doc, target):
        p = target.ownerDocument.createElement('w:p')
        if self.style is not None:
            pPr = target.ownerDocument.createElement('w:pPr')
            pStyle = target.ownerDocument.createElement('w:pStyle')
            pStyle.setAttribute('w:val', doc.styles[self.style])
            pPr.appendChild(pStyle)
            p.appendChild(pPr)
        Text('%s ' % self.name).append_to(doc, p)
        self.counter.append_to(doc, p)
        Text(': ').append_to(doc, p)
        for c in self.content:
            c.append_to(doc, p)
        target.appendChild(p)

class Table(object):
    '''Simple table.
    
    cells should be a 2dim array.
    '''
    def __init__(self, cells, caption = None, style = None):
        self.cells = cells
        if caption is not None:
            self.caption = Caption('Table', caption)
        else:
            self.caption = None
        self.style = style

    def append_to(self, doc, target):
        if self.caption is not None:
            self.caption.append_to(doc, target)
        tbl = target.ownerDocument.createElement('w:tbl')
        tblPr = target.ownerDocument.createElement('w:tblPr')
        tblW = target.ownerDocument.createElement('w:tblW')
        tblW.setAttribute('w:w', '5000')
        tblW.setAttribute('w:type', 'pct')
        tblPr.appendChild(tblW)
        if self.style is not None:
            tblStyle = target.ownerDocument.createElement('w:tblStyle')
            tblStyle.setAttribute('w:val', self.style)
            tblPr.appendChild(tblStyle)
        else:
            tblBorders = target.ownerDocument.createElement('w:tblBorders')
            bottom = target.ownerDocument.createElement('w:bottom')
            bottom.setAttribute('w:val', 'single')
            bottom.setAttribute('w:sz', '4')
            bottom.setAttribute('wx:bdrwidth', '10')
            bottom.setAttribute('w:space', '0')
            bottom.setAttribute('w:color', 'auto')
            tblBorders.appendChild(bottom)
            tblPr.appendChild(tblBorders)
        tblLook = target.ownerDocument.createElement('w:tblLook')
        tblLook.setAttribute('w:val', '01E0')
        tblPr.appendChild(tblLook)
        tbl.appendChild(tblPr)
        tblGrid = target.ownerDocument.createElement('w:tblGrid')
        for c in self.cells[0]:
            tblGrid.appendChild(target.ownerDocument.createElement('w:gridCol'))
        tbl.appendChild(tblGrid)
        target.appendChild(tbl)
        for row in self.cells:
            self._append_row_to(doc, tbl, row)

    def _append_row_to(self, doc, target, row):
        tr = target.ownerDocument.createElement('w:tr')
        for c in row:
            tc = target.ownerDocument.createElement('w:tc')
            if isinstance(c, Paragraph):
                c.append_to(doc, tc)
            elif isinstance(c, (str, unicode, Text)):
                Paragraph(c).append_to(doc, tc)
            else:
                Paragraph(c).append_to(doc, str(tc))
            tr.appendChild(tc)
        target.appendChild(tr)

class Figure(object):
    '''External image.
    '''
    def __init__(self, fname, size, caption = None):
        self.fname = fname
        self.size = size
        if caption is not None:
            self.caption = Caption('Figure', caption)
        else:
            self.caption = None

    def append_to(self, doc, target):
        name = self._copy_media(doc)
        media_id = doc.appendMedia(name)
        p = target.ownerDocument.createElement('w:p')
        r = target.ownerDocument.createElement('w:r')
        drawing = target.ownerDocument.createElement('w:drawing')
        inline = target.ownerDocument.createElement('wp:inline')
        extent = target.ownerDocument.createElement('wp:extent')
        extent.setAttribute('cx', '%.0f' % (self.size[0] * 911400))
        extent.setAttribute('cy', '%.0f' % (self.size[1] * 911400))
        inline.appendChild(extent)
        docpr = target.ownerDocument.createElement('wp:docPr')
        docpr.setAttribute('id', '1')
        docpr.setAttribute('descr', 'Grafik 0')
        docpr.setAttribute('name', name)
        inline.appendChild(docpr)
        cNvGraphicFramePr = target.ownerDocument.createElement('wp:cNvGraphicFramePr')
        graphicFrameLocks = target.ownerDocument.createElement('a:graphicFrameLocks')
        graphicFrameLocks.setAttribute('noChangeAspect', '1')
        cNvGraphicFramePr.appendChild(graphicFrameLocks)
        inline.appendChild(cNvGraphicFramePr)
        graphic = target.ownerDocument.createElement('a:graphic')
        graphicdata = target.ownerDocument.createElement('a:graphicData')
        graphicdata.setAttribute('uri', 'http://schemas.openxmlformats.org/drawingml/2006/picture')
        pic = target.ownerDocument.createElement('pic:pic')
        pic.setAttribute('xmlns:pic', ns['pic'])
        nvpicpr = target.ownerDocument.createElement('pic:nvPicPr')
        cnvpr = target.ownerDocument.createElement('pic:cNvPr')
        cnvpr.setAttribute('id', '1')
        cnvpr.setAttribute('name', name)
        nvpicpr.appendChild(cnvpr)
        cnvpicpr = target.ownerDocument.createElement('pic:cNvPicPr')
        nvpicpr.appendChild(cnvpicpr)
        pic.appendChild(nvpicpr)
        blipFill = target.ownerDocument.createElement('pic:blipFill')
        blip = target.ownerDocument.createElement('a:blip')
        blip.setAttribute('r:embed', media_id)
        blip.setAttribute('cstate', 'print')
        blipFill.appendChild(blip)
        stretch = target.ownerDocument.createElement('a:stretch')
        fillRect = target.ownerDocument.createElement('a:fillRect')
        stretch.appendChild(fillRect)
        blipFill.appendChild(stretch)
        pic.appendChild(blipFill)
        sppr = target.ownerDocument.createElement('pic:spPr')
        xfrm = target.ownerDocument.createElement('a:xfrm')
        off = target.ownerDocument.createElement('a:off')
        off.setAttribute('x', '0')
        off.setAttribute('y', '0')
        xfrm.appendChild(off)
        ext = target.ownerDocument.createElement('a:ext')
        ext.setAttribute('cx', '%.0f' % (self.size[0] * 911400))
        ext.setAttribute('cy', '%.0f' % (self.size[1] * 911400))
        xfrm.appendChild(ext)
        sppr.appendChild(xfrm)
        prstGeom = target.ownerDocument.createElement('a:prstGeom')
        prstGeom.setAttribute('prst', 'rect')
        sppr.appendChild(prstGeom)
        pic.appendChild(sppr)
        graphicdata.appendChild(pic)
        graphic.appendChild(graphicdata)
        inline.appendChild(graphic)
        drawing.appendChild(inline)
        r.appendChild(drawing)
        p.appendChild(r)
        target.appendChild(p)
        if self.caption is not None:
            self.caption.append_to(doc, target)

    def _copy_media(self, doc):
        newname = os.path.basename(self.fname)
        shutil.copy(self.figure, os.path.join(doc.mediadir, newname))
        return newname

class MatplotlibFigure(Figure):
    '''Image made from a matplotlib figure
    '''
    def __init__(self, fig, caption = None):
        Figure.__init__(self, 'matplotlib', fig.get_size_inches(), caption)
        self.fig = fig

    def _copy_media(self, doc):
        FigureCanvasAgg(self.fig)
        tmpf = tempfile.mkstemp(dir = doc.mediadir, prefix = 'image',
                                suffix = '.png')
        os.close(tmpf[0])
        self.fig.savefig(tmpf[1], format = 'png', dpi=300)
        return os.path.basename(tmpf[1])

class PageBreak(object):
    '''Page break
    '''
    def append_to(self, doc, target):
        p = target.ownerDocument.createElement('w:p')
        r = target.ownerDocument.createElement('w:r')
        br = target.ownerDocument.createElement('w:br')
        br.setAttribute('w:type', 'page')
        r.appendChild(br)
        p.appendChild(r)
        target.appendChild(p)

class List(object):
    '''(Unnumbered) List
    '''

    bullets = [ u'●', u'○', '-', u'•', u'◦', u'-', u'▪', u'▫', u'-' ]

    def __init__(self, rows = None, style = None, align = None, format = None):
        self.style = style
        self.align = align
        self.format = format
        self.indent = 360
        self.hanging = 360
        self.rows = [ ]
        if rows:
            for r in rows:
                self.__iadd__(r)

    def __iadd__(self, row):
        if isinstance(row, (List, Paragraph)):
            self.rows.append(row)
        elif isinstance(row, list):
            pr = list()
            for r in row:
                if isinstance(r, (List, Paragraph)):
                    pr.append(r)
                else:
                    pr.append(Paragraph(r, style = self.style, 
                                        align = self.align))
            self.rows.append(pr)
        else:
            self.rows.append(Paragraph(row, style = self.style, 
                                       align = self.align))

    def append_to(self, doc, target, parent = None, indent = 0):
        indent += self.indent
        level = parent[0] + 1 if parent else 0
        format = self.format or List.bullets[level]
        numId = doc.numbering(parent, indent, self.hanging, format)
        for row in self.rows:
            if isinstance(row, List):
                row.append_to(doc, target, numId, indent = indent)
            elif isinstance(row, list):
                row[0].append_to(doc, target, numbering = numId, 
                                 indent = indent)
                for r in row[1:]:
                    r.append_to(doc, target, indent = indent)
            else:
                row.append_to(doc, target, numbering = numId, indent = indent)

class Numbering(object):
    def __init__(self, sfile = None):
        self.nums = dict()
        self.maxnumber = 0
        if sfile is not None:
            self.doc = minidom.parse(sfile)
            for n in self.doc.getElementsByTagName('w:num'):
                num = int(n.getAttribute('w:numId'))
                self.maxnumber = max(self.maxnumber, num)
            for an in self.doc.getElementsByTagName('w:abstractNum'):
                num = int(n.getAttribute('w:abstractNumId'))
                self.maxnumber = max(self.maxnumber, num + 1)
        else:
            self.doc = minidom.Document()
            n = self.doc.createElement('w:numbering')
            n.setAttribute('xmlns:w', ns['w'])
            self.doc.appendChild(n)

    def _get_format(self, content, level = 0):
        for cstart, c in enumerate(content):
            if c.isalnum():
                break
        for clen, c in enumerate(content[cstart:]):
            if not c.isalnum():
                break
        prefix = content[:cstart]
        suffix = content[cstart+clen:]
        c = content[cstart:cstart+clen]
        if not c:
            start = None
            fmt = 'bullet'
            txt = content
        elif c.isdigit():
            fmt = 'decimal'
            start = int(c)
            txt = '%s%%%d%s' % (prefix, level + 1, suffix)
        else:
            formats = { 'a':'lowerLetter', 
                        'A':'upperLetter',
                        'i':'lowerRoman',
                        'I':'upperRoman',
                        }
            fmt = formats.get(c)
            start = 1
            txt = '%s%%%d%s' % (prefix, level + 1, suffix)
        return start, fmt, txt

    def _new_numbering(self):
        an = self.doc.createElement('w:abstractNum')
        an.setAttribute('w:abstractNumId', '%i' % self.maxnumber)
        nums = self.doc.getElementsByTagName('w:num')
        rootElement = self.doc.getElementsByTagName('w:numbering')[0]
        if nums:
            rootElement.insertBefore(an, nums[0])
        else:
            rootElement.appendChild(an)
        hml = self.doc.createElement('w:multiLevelType')
        hml.setAttribute('w:val', 'hybridMultilevel')
        an.appendChild(hml)
        num = self.doc.createElement('w:num')
        num.setAttribute('w:numId', '%i' % (self.maxnumber + 1))
        ani = self.doc.createElement('w:abstractNumId')
        ani.setAttribute('w:val', '%i' % self.maxnumber)
        num.appendChild(ani)
        rootElement.appendChild(num)
        self.maxnumber += 1
        self.nums[self.maxnumber] = { 0:an }
        return an, self.maxnumber

    def add(self, parent, indent, hanging, content):
        n = self.nums.get(parent[1]) if parent else None
        if parent and n and not (parent[0] + 1) in n:
            level = parent[0] + 1
            an = n.get(parent[0])
            numId = parent[1]
            self.nums[numId][level] = an 
        else:
            level = 0
            an, numId = self._new_numbering()

        lvl = self.doc.createElement('w:lvl')
        lvl.setAttribute('w:ilvl', '%i' % level)
        start, fmt, txt = self._get_format(content, level)
        if start:
            st = self.doc.createElement('w:start')
            st.setAttribute('w:val', '%i' % start)
            lvl.appendChild(st)
        if fmt:
            numfmt = self.doc.createElement('w:numFmt')
            numfmt.setAttribute('w:val', fmt)
            lvl.appendChild(numfmt)
        if txt:
            lvltext = self.doc.createElement('w:lvlText')
            lvltext.setAttribute('w:val', txt)
            lvl.appendChild(lvltext)
        if indent or hanging:
            pPr = self.doc.createElement('w:pPr')
            ind = self.doc.createElement('w:ind')
            if indent:
                ind.setAttribute('w:left', '%i' % indent)
            if hanging:
                ind.setAttribute('w:hanging', '%i' % hanging)
            pPr.appendChild(ind)
            lvl.appendChild(pPr)
        an.appendChild(lvl)
        return (level, numId)
