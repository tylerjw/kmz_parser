"""
Parser GUI application: Parses KMZ files and outputs to text box or csv files.

Author: Tyler Weaver

Dependencies:
Python v2.6.5
CoordConverter.py (v0.1) - library for coordinate conversions

Versions:
0.1: input and output and button
0.2: output text scrollbox, comma or tab separation
0.3: Browse popup options
0.4: file output checkbox, error handling of input file
0.5: fixed compatabiliity with Palantir output with points
0.6(06/07/13): renamed file to remove annoying command window
"""

from Tkinter import *
import tkFileDialog
from ScrolledText import *
import xml.sax, xml.sax.handler
from CoordConverter import CoordTranslator
import re
from zipfile import ZipFile, BadZipfile

version = "0.6"

class PlacemarkHandler(xml.sax.handler.ContentHandler):
    def __init__(self):
        self.inName = False # handle XML parser events
        self.inPlacemark = False
        self.mapping = {} # a state machine model
        self.buffer = ""
        self.name_tag = ""
        
    def startElement(self, name, attributes):
        if name == "Placemark": # on start Placemark tag
            self.inPlacemark = True
            self.buffer = "" 

        if self.inPlacemark:
            if name == "name": # on start title tag
                self.inName = True # save name text to follow
            
    def characters(self, data):
        if self.inPlacemark: # on text within tag
            self.buffer += data # save text if in title
            

    def endElement(self, name):
        self.buffer = self.buffer.strip('\n\t')
        
        if name == "Placemark":
            self.inPlacemark = False
            self.name_tag = "" #clear current name
        
        elif name == "name" and self.inPlacemark:
            self.inName = False # on end title tag            
            self.name_tag = self.buffer.strip()
            self.mapping[self.name_tag] = {}

        elif self.inPlacemark:
            if name in self.mapping[self.name_tag]:
                self.mapping[self.name_tag][name] += self.buffer
            else:
                self.mapping[self.name_tag][name] = self.buffer

        self.buffer = ""

def build_table(mapping,delin=1):
    sep = ','
    if delin == 2:
        sep = '\t'
        
    output = 'Name'+sep+'Coordinates\n'
    points = ''
    lines = ''
    shapes = ''
    for key in mapping:
        mgrs_str = ''
        for coord in mapping[key]['mgrs']:
            mgrs_str += coord + sep
        if 'LookAt' in mapping[key]:
            points += key +sep+ mgrs_str + "\n"
        elif 'LineString' in mapping[key]:
            lines += key +sep+ mgrs_str + "\n"
        else:
            shapes += key +sep+ mgrs_str + "\n"

    output += points + lines + shapes
    return output

def coordinates_to_mgrs(mapping):
    c = CoordTranslator()
    re_latlong = re.compile('[+-]?[\d.]+,[+-]?[\d.]+[, 0]{0,2}')
    for key in mapping:
        output = []
        
        for item in re_latlong.finditer(mapping[key]['coordinates']):
            coord = item.group()
            span = re.compile('[+-]?[\d.]+,').search(coord).span()
            lon = float(coord[span[0]:span[1]-1])
            span = re.compile('[+-]?[\d.]+').search(coord,span[1]).span()
            lat = float(coord[span[0]:span[1]-1])
            #print [lat, lon]
            #print c.AsMGRS([lat,lon])
            output.append(c.AsMGRS([lat,lon], spaces = False))

        mapping[key]['mgrs'] = output

    return mapping

class ParserWindow(Frame):
    def __init__(self):
        Frame.__init__(self)
        self.pack(expand=YES, fill=BOTH)
        self.master.title('KMZ to CSV Parser v' + version)
        self.master.iconname('KMZ Parser')

        self.in_filename = StringVar()
        self.out_filename = StringVar()

        self.in_filename.set('test.kmz')
        self.out_filename.set('output.csv')

        row1 = Frame(self)
        Label(row1, text="Input:", width=10).pack(side=LEFT, pady=10)
        self.in_ent = Entry(row1, width=80, textvariable=self.in_filename)
        self.in_ent.pack(side=LEFT)
        self.in_ent.bind('<Key-Return>', self.Parse)
        Button(row1, text="Browse", width=10, command=self.get_kmz_filename).pack(side=LEFT, padx=5)
        row1.pack(side=TOP, ipadx=15)

        row2 = Frame(self)
        Label(row2, text="Output: ", width=10).pack(side=LEFT, pady=5)
        self.out_ent = Entry(row2, width=80, textvariable=self.out_filename)
        self.out_ent.pack(side=LEFT)
        self.out_ent.bind('<Key-Return>', self.Parse)
        Button(row2, text="Browse", width=10, command=self.get_output_filename).pack(side=LEFT, padx=5)
        row2.pack(side=TOP, ipadx=15)

        self.sepvar = IntVar()
        ent3 = Frame(self)
        ent4 = Frame(ent3)

        self.en_file_out = IntVar()
        r1 = Radiobutton(ent4, text="Comma Separated", value=1, variable=self.sepvar, command=self.Parse)
        r2 = Radiobutton(ent4, text="Tab Separated", value=2, variable=self.sepvar, command=self.Parse)
        chb = Checkbutton(ent4, text="File Output", variable=self.en_file_out, command=self.Parse)
        chb.deselect()
        self.sepvar.set(1)
        r1.pack(side=LEFT)
        r2.pack(side=LEFT)
        chb.pack(side=LEFT)
        
        btn = Button(ent3, text="Parse", command=self.Parse, width=15)
        btn.pack(side=LEFT, padx=5, pady=10)

        ent4.pack(side=LEFT)
        ent3.pack(side=TOP)
        
        self.scrolled_text = ScrolledText(self, height=15, state="disabled",
                                          padx=10, pady=10,
                                          wrap="word")
        self.scrolled_text.pack(side=LEFT, fill=BOTH, expand=1, padx=5, pady=5)

    def get_kmz_filename(self):
        filename = tkFileDialog.askopenfilename(parent=self,
                                                filetypes=[('All Files', '*'),
                                                           ('KMZ File', '*.kmz')],
                                                multiple=False)
        self.in_filename.set(filename)

    def get_output_filename(self):
        filename = tkFileDialog.asksaveasfilename(parent=self,
                                                  filetypes=[('All Files', '*'),
                                                           ('Text File', '*.txt'),
                                                             ('CSV File', '*.csv')],
                                                default='output.csv')
        self.out_filename.set(filename)
               
    def Parse(self, event=None):
        try:
            kmz = ZipFile(self.in_filename.get(),'r')
        except IOError:
            self.scrolled_text['state'] = "normal"
            self.scrolled_text.delete(1.0, END)
            self.scrolled_text.insert(END,"Error Opening File!  Check input filename.")
            self.scrolled_text['state'] = "disabled"
            return
        except BadZipfile:
            self.scrolled_text['state'] = "normal"
            self.scrolled_text.delete(1.0, END)
            self.scrolled_text.insert(END,"Bad ZIP file!  Check that input is a .KMZ file.")
            self.scrolled_text['state'] = "disabled"
            return
        
        kml = kmz.open('doc.kml','r')
        parser = xml.sax.make_parser()
        handler = PlacemarkHandler()
        parser.setContentHandler(handler)
        parser.parse(kml)
        kmz.close()
        handler.mapping = coordinates_to_mgrs(handler.mapping)
        outstr = build_table(handler.mapping, self.sepvar.get())
        if self.en_file_out.get():
            f = open(self.out_filename.get(), "w")
            f.write(outstr)
            f.close()
            outstr = " -- Saved to %s -- \n" %self.out_filename.get() + outstr
            
        self.scrolled_text['state'] = "normal"
        self.scrolled_text.delete(1.0, END)
        self.scrolled_text.insert(END,outstr)
        self.scrolled_text['state'] = "disabled"
        
if __name__ == '__main__':
    ParserWindow().mainloop()
