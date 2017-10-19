from sys import argv
from lxml import etree, objectify
from lxml.builder import E
from lxml.builder import ElementMaker
import argparse
import os.path
import mammoth
import zipfile

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

# an empty dict for our ultimate parsed data
verboseAttrs = {}

# function to read the styles.xml file from within the docx;
# this is used for extracting style names and formatting information.
def getWordStyles(file):
  zip = zipfile.ZipFile(file)
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
def getAllStyles(file):
  # parse the incoming XML
  source = getWordStyles(file)
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

# add the formatting info back to the HTML as attributes on each element
def addAttrs(html, myDict):
  root = etree.HTML(html)
  for style, vals in myDict.iteritems():
    for para in root.findall(".//p[@class='" + style + "']"):
      for key, val in vals.iteritems():
        para.attrib[key] = val
  newHTML = etree.tostring(root, encoding="UTF-8", standalone=True, xml_declaration=True)
  return newHTML

def sanitizeHTML(html):
  root = etree.HTML(html)
  newHTML = etree.tostring(root, encoding="UTF-8", standalone=True, xml_declaration=True)
  return newHTML

verboseAttrs = getAllStyles(docxfile)

# create the style map if requested
if args.mapStyles == True:
  style_map = '"""'
  for style, vals in verboseAttrs.iteritems():
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
result = mammoth.convert_to_html(docxfile, style_map=style_map)
html = result.value
messages = result.messages

# add the verbose attributes to the output HTML if requested
if args.preserveFormatting == True:
  html = addAttrs(html, verboseAttrs)
else:
  html = sanitizeHTML(html)

# write to a new HTML document
output = open('output.html', 'w')
output.write(html)
output.close()