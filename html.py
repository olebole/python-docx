import codecs
import os
import tempfile
import shutil
from xml.dom import minidom

try:
    import matplotlib.figure
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    _have_matplotlib = True
except:
    _have_matplotlib = False

class Document(object):
    def __init__(self, fname = None, mode = 'copyonwrite'):
        self.fname = fname
        self.mode = mode
        self.tmpdir = tempfile.mkdtemp(prefix = 'html')
        self.mediadir = os.path.join(self.tmpdir, 'images')
        self.property = dict()
        os.mkdir(self.mediadir)

        if fname is None or mode is 'create':
            self._createdefault()
        elif os.path.exists(fname):
            self._load(fname)
        elif mode is 'create' or mode is 'append':
            self._createdefault()
        else:
            raise IOError('%s does not exist' % fname)


    def _createdefault(self):
        self.styles = {
            'td': {'border':'1px solid',
                   'padding':'2px'},
            'table': {'border-spacing':'0'},
            'div.figure': {'width':'85%' },
            'div.table': {'width':'85%' },
            }
        doc = minidom.Document()
        wdoc = doc.createElement('html')
        doc.appendChild(wdoc)
        self.header = doc.createElement('head')
        v = ''
        for name, style in self.styles.items():
            s0 = '\n'
            for p, s in style.items():
                s0 += '  %s: %s;\n' % (p,s)
            v += '%s {\n%s}\n' % (name, s0)
        st = doc.createElement('style')
        st.setAttribute('type', 'text/css')
        st.appendChild(doc.createTextNode(v))
        self.header.appendChild(st)
        wdoc.appendChild(self.header)
        self.body = doc.createElement('body')
        wdoc.appendChild(self.body)

    def _load(self, filename):
        wdoc = minidom.parse(open(filename)).documentElement
        self.body = wdoc.getElementsByTagName('body')[0]
        self.header = wdoc.getElementsByTagName('head')[0]

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

    def writeto(self, filename):
        fp = codecs.open(filename, 'w', 'utf-8')
        for key, value in self.property.items():
            st = self.header.ownerDocument.createElement('meta')
            st.setAttribute('name', key)
            st.setAttribute('content', value)
            self.header.appendChild(st)
        self.body.ownerDocument.writexml(fp, encoding = "UTF-8")
#        self.body.ownerDocument.writexml(fp, "", "  ", "\n", encoding = "UTF-8")
        images = os.path.join(os.path.dirname(filename), 
                              os.path.basename(self.mediadir))
        if not os.path.exists(images):
            os.mkdir(images)
        for f in os.listdir(self.mediadir):
            shutil.copy(os.path.join(self.mediadir, f), images)
        fp.close()

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
        p = target.ownerDocument.createElement('span')
        style = ''
        if self.bold is not None:
            style += 'font-weight: %s; ' % ('bold' if self.bold else 'normal')
        if self.italic is not None:
            style += 'font-style: %s; ' % ('italic' if self.italic else 'normal')
        if self.underline is not None:
            styles = {
                0:'none', False:'none',
                1:'underline', True:'underline', '_':'underline',
                2:'underline', '=':'underline',
                '#':'underline',
                '-':'underline',
                ',':'underline',
                '.':'underline',
                ';':'underline',
                }
            style += 'text-decoration: %s;' % styles[self.styles]
        if style:
            p.setAttribute('style', style)
        p.appendChild(target.ownerDocument.createTextNode(self.content.strip()))
        target.appendChild(p)

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
        self.content.append(other if isinstance(other, (Text, List)) else Text(other))
        return self

    def append_to(self, doc, target):
        p = target.ownerDocument.createElement(self.style or 'p')
        styles = ''
        alignments = {
            'l':'left', 'left':'left', '<':'left',
            'r':'right', 'right':'right', '>':'right',
            'c':'center', 'center':'center', 
            'b':'both', 'block':'both', 'both':'both', '=':'both',
            }
        a = alignments.get(self.align)
        if a:
            styles += 'alignment:%s; ' % a
        if styles:
            p.setAttribute('style', styles)
        for c in self.content:
            c.append_to(doc, p)
        target.appendChild(p)

class Header(Paragraph):
    '''Caption header.
    '''
    def __init__(self, level, content):
        Paragraph.__init__(self, content, style = 'h%i' % level)

class List(object):
    '''(Unnumbered) List
    '''
    def __init__(self, rows = None, align = None, format = None):
        self.format = format
        self.align = align
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
                    pr.append(Paragraph(r, align = self.align))
            self.rows.append(pr)
        else:
            self.rows.append(Paragraph(row, align = self.align))

    def append_to(self, doc, target):
        a = self.format[0] if self.format else None
        if a in ('1', 'a', 'A', 'i', 'I'):
            p = target.ownerDocument.createElement('ol')
        else:
            p = target.ownerDocument.createElement('ul')
        for row in self.rows:
            li = target.ownerDocument.createElement('li')
            row.append_to(doc, li)
            p.appendChild(li)
        target.appendChild(p)

class Caption(Paragraph):
    '''Table or figure caption
    '''
    def __init__(self, name, content):
        Paragraph.__init__(self, content)
        self.content.insert(0, Text(name + ': ', bold=True))

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
        d = target.ownerDocument.createElement('div')
        d.setAttribute('class', 'table')
        if self.caption is not None:
            self.caption.append_to(doc, d)
        tbl = target.ownerDocument.createElement('table')
        for row in self.cells:
            self._append_row_to(doc, tbl, row)
        d.appendChild(tbl)
        target.appendChild(d)

    def _append_row_to(self, doc, target, row):
        tr = target.ownerDocument.createElement('tr')
        for c in row:
            tc = target.ownerDocument.createElement('td')
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
        d = target.ownerDocument.createElement('div')
        d.setAttribute('class', 'figure')
        i = target.ownerDocument.createElement('img')
        i.setAttribute('src', os.path.join(os.path.basename(doc.mediadir), 
                                           name))
        i.setAttribute('width', '100%')
        d.appendChild(i)
        if self.caption is not None:
            self.caption.append_to(doc, d)
        target.appendChild(d)

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
        p = target.ownerDocument.createElement('hline')
        target.appendChild(p)
