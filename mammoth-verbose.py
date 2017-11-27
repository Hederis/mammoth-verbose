from sys import argv
from lxml import etree, objectify
from lxml.builder import E
from lxml.builder import ElementMaker
from docx import Document
from docx.shared import Inches, Pt
from docx.text.run import Font, Run
import shutil
import argparse
import os.path
import mammoth
import zipfile
import inspect

# function to check if the input file is valid
def is_valid_file(parser, arg):
  if not os.path.exists(arg):
    parser.error("The file %s does not exist!" % arg)
  else:
    return open(arg, 'r')  # return an open file handle

# defining the program options
parser = argparse.ArgumentParser(description='While using the Mammoth docx converter, add options to preserve source class names and formatting information as attributes.')
parser.add_argument("-i", dest="filename", required=True,
                    help="The name of the file to read.", metavar="FILE",
                    type=lambda x: is_valid_file(parser, x))
parser.add_argument('--map', dest='mapStyles', action='store_true', default=True,
                   help='Create a custom map to preserve the source docx style names as classes in the output HTML. Default is True.')
parser.add_argument('--verbose', dest='preserveFormatting', action='store_true', default=False,
                   help='Preserve any formatting applied to the docx styles as attributes in the output HTML. Default is False.')

args = parser.parse_args()

docxfile = args.filename
fileName = docxfile.name

# an empty dict for our ultimate parsed data
verboseAttrs = {}

# function to read the styles.xml file from within the docx;
# this is used for extracting style names and formatting information.
def getWordStyles(myfile):
  fileName = myfile.name
  filePath = os.path.splitext(myfile.name)[0]
  newName = filePath + ".zip"
  shutil.copyfile(fileName, newName)
  zip = zipfile.ZipFile(newName)
  xml_content = zip.read('word/styles.xml')
  return xml_content

# get all the formatting attributes for a style
def getAttrs(element, inputKey="data", inputVal="", inputDict={}):
  attrKey = element.tag
  attrKey = inputKey + "-" + attrKey.split("}").pop()
  attributes = element.attrib
  children = list(element)
  attrVal = ""
  if len(attributes) == 0 and len(children) == 0:
    attrVal = "true"
  elif len(attributes) == 1 and element.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"):
    attrVal = element.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
  elif len(attributes) > 0:
    for name, value in sorted(element.items()):
      name = name.split("}").pop()
      attrVal = attrVal + "%s:%r;" % (name, value)
  # loop through sub-children and add as attr
  walkChildren(element, attrKey, "", inputDict)

  return attrKey, attrVal

# walk through the element tree for a style
def walkChildren(element, inputKey="data", inputVal="", inputDict={}):
  children = list(element)
  attrKey = ""
  attrVal = ""
  for child in element:
    attrKey, attrVal = getAttrs(child, inputKey, "", inputDict)
    inputDict[attrKey] = attrVal
  return attrKey, attrVal, inputDict

# bringing together getWordStyles, getAttrs, and walkChildren
# to create the final verboseDict of all style names and their
# formatting information.
def getAllStyles(myfile):
  # parse the incoming XML
  source = getWordStyles(myfile)
  root = etree.fromstring(source)

  allStyles = {}
  # get all paragraph styles
  for style in root.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}style[@{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type='paragraph']") :
    styleID = style.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId")
    # reset the dictionary for this style's children
    allAttr = {}
    # get all child elements of the style
    children = list(style)
    # walk the child tree to collect all elements and their attributes into the dictionary
    attrKey, attrVal, allAttr = walkChildren(style, "data", "", allAttr)
    allAttr['data-w-type'] = 'p'
    # add the style ID to the master dictionary with value = the collected child elements
    allStyles[styleID] = allAttr
        
  # get all character styles
  for style in root.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}style[@{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type='character']") :
    styleID = style.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId")
    # reset the dictionary for this style's children
    allAttr = {}
    # get all child elements of the style
    children = list(style)
    # walk the child tree to collect all elements and their attributes into the dictionary
    attrKey, attrVal, allAttr = walkChildren(style, "data", "", allAttr)
    allAttr['data-w-type'] = 'r'
    # add the style ID to the master dictionary with value = the collected child elements
    allStyles[styleID] = allAttr

  return allStyles

def getAttr(self):
  for attr in self:
      value = getattr(format, attr, None)
      if value != None:
        print(attr, value)

def getDirectFormatting(myfile):
  fobj = open(myfile,'rb')
  document = Document(fobj)
  paragraphs = document.paragraphs
  parastyles = {}
  for para in paragraphs:
    # Get paragraph formatting
    format = para.paragraph_format
    sublist = [a for a in dir(format) if not a.startswith('_') and a != 'element' and a != 'tab_stops']
    parastyles['alignment'] = para.paragraph_format.alignment
    #Element is assumed to be w:p
    #parastyles['element'] = para.paragraph_format.element
    parastyles['first_line_indent'] = para.paragraph_format.first_line_indent
    parastyles['keep_together'] = para.paragraph_format.keep_together
    parastyles['keep_with_next'] = para.paragraph_format.keep_with_next
    parastyles['left_indent'] = para.paragraph_format.left_indent
    parastyles['line_spacing'] = para.paragraph_format.line_spacing
    parastyles['line_spacing_rule'] = para.paragraph_format.line_spacing_rule
    parastyles['page_break_before'] = para.paragraph_format.page_break_before
    #parastyles['part'] = para.paragraph_format.part
    parastyles['right_indent'] = para.paragraph_format.right_indent
    parastyles['space_after'] = para.paragraph_format.space_after
    parastyles['space_before'] = para.paragraph_format.space_before
    #We'll ignore tab stops for now
    #parastyles['tab_stops'] = para.paragraph_format.tab_stops
    parastyles['widow_control'] = para.paragraph_format.widow_control
    for key,val in parastyles.items():
      if val != None:
        print(key, ": ", val)
        # TO DO: do something with paragraph formatting
    # for attr in sublist:
    #   selector = "format." + attr
    #   val = selector
    #   print(attr, ": ", val)
    #print(sublist)
    # for attr in sublist:
    #   val = format[attr]
    #   print(attr + ": " + val)
      # add to new style def
    runs = para.runs
    charstyles = {}
    for run in runs:
      font = run.font
      sublist = [a for a in dir(font) if not a.startswith('_') and a != 'element']
      charstyles['all_caps'] = run.font.all_caps
      charstyles['bold'] = run.font.bold
      charstyles['color'] = run.font.color.rgb
      charstyles['complex_script'] = run.font.complex_script
      charstyles['cs_bold'] = run.font.cs_bold
      charstyles['cs_italic'] = run.font.cs_italic
      charstyles['double_strike'] = run.font.double_strike
      charstyles['emboss'] = run.font.emboss
      charstyles['hidden'] = run.font.hidden
      charstyles['highlight_color'] = run.font.highlight_color
      charstyles['imprint'] = run.font.imprint
      charstyles['italic'] = run.font.italic
      charstyles['math'] = run.font.math
      charstyles['name'] = run.font.name
      charstyles['no_proof'] = run.font.no_proof
      charstyles['outline'] = run.font.outline
      charstyles['part'] = run.font.part
      charstyles['rtl'] = run.font.rtl
      charstyles['shadow'] = run.font.shadow
      charstyles['size'] = run.font.size
      charstyles['small_caps'] = run.font.small_caps
      charstyles['snap_to_grid'] = run.font.snap_to_grid
      charstyles['spec_vanish'] = run.font.spec_vanish
      charstyles['strike'] = run.font.strike
      charstyles['subscript'] = run.font.subscript
      charstyles['superscript'] = run.font.superscript
      charstyles['underline'] = run.font.underline
      charstyles['web_hidden'] = run.font.web_hidden
      print(sublist)
      for attr in sublist:
        value = getattr(font, attr, None)
        if value != None:
          pass
          #print(attr, ": ", value)
      # for key, val in font.items():
      #   print(key + ": " + val)

# add the formatting info back to the HTML as attributes on each element
def addAttrs(html, myDict):
  root = etree.HTML(html)
  for style, vals in myDict.items():
    for para in root.findall(".//p[@class='" + style + "']"):
      for key, val in vals.items():
        para.attrib[key] = val
  newHTML = etree.tostring(root, encoding="UTF-8", standalone=True, xml_declaration=True)
  return newHTML

def sanitizeHTML(html):
  root = etree.HTML(html)
  newHTML = etree.tostring(root, encoding="UTF-8", standalone=True, xml_declaration=True)
  return newHTML

getDirectFormatting(fileName)

verboseAttrs = getAllStyles(docxfile)

# create the style map if requested
if args.mapStyles == True:
  style_map = '"""'
  for style, vals in verboseAttrs.items():
    sourceName = vals['data-name']
    destName = style
    # mapping paragraphs
    if vals['data-w-type'] == 'p':
      thisMap = "p[style-name='" + sourceName + "'] => p." + destName + ":fresh"
    # mapping runs
    else:
      thisMap = "r[style-name='" + sourceName + "'] => span." + destName
    # write this map to the map file
    style_map = style_map + "\n" + thisMap

  style_map = style_map + '\n"""'

# convert with mammoth
fobj = open(fileName,'rb')
result = mammoth.convert_to_html(fobj, style_map=style_map)
html = result.value
messages = result.messages

# add the verbose attributes to the output HTML if requested
if args.preserveFormatting == True:
  html = addAttrs(html, verboseAttrs)
else:
  html = sanitizeHTML(html)

# write to a new HTML document
output = open('output.html', 'w')
output.write(str(html))
output.close()